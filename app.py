"""
GC-MS FID 成分分類網站
自動對 Shimadzu GC-MS 輸出的 xlsx 執行成分分類、顏色標記、統計加總
"""

import io
import os
import re
import json
import time
import base64
import pandas as pd
from flask import Flask, request, send_file, render_template, jsonify
from werkzeug.exceptions import HTTPException
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

COLOR_MAP = {
    'n-alkane':     'FF8C00',
    'iso-alkane':   'FF6600',
    'cyclic':       'FFD700',
    'Alkene':       'ADD8E6',
    'Aromatic':     '92D050',
    'Polyaromatic': 'C5B0D5',
    'others':       'FA8072',
    'solvent':      'BFBFBF',
}

# ── 碳數視窗邊界（預設值，可由前端表單覆寫）────────────────────────
# 依滯留時間切 C8- / C8-16 / C16+ 三段。邊界應該是「該次分析 n-alkane 標記峰
# 的實際滯留時間」，隨管柱、溫度程式改變，所以做成可覆寫的預設值而不是常數。
RT_C8_MINUS_MAX = 2.178
RT_C816_MAX     = 17.6

# ── 溶劑峰 ─────────────────────────────────────────────────────────
# 樣品進樣前通常用 n-hexane 之類的溶劑稀釋，溶劑峰在層析圖最前端、面積極大，
# 而且因為嚴重過載而拖尾、被積分成好幾個峰；NIST 對過載峰的比對品質很差，
# 常回報成同分異構物（R5 實測兩個溶劑峰都被命名為 "PENTANE, 3-METHYL-"）。
# 溶劑不是產物，必須整筆排除，否則濃度分母是錯的。
SOLVENT_MIN_CONC = 2.0        # 溶劑峰的單峰 Conc% 門檻
SOLVENT_CATEGORY = 'solvent'  # 被排除的峰在 Category 欄的標記

# 過載的 n-hexane 溶劑峰，NIST 會回報成這幾個 C6H14 同分異構物之一。
# 單看濃度門檻不夠：R10 的溶劑峰只有 0.91%（樣品濃度高、溶劑相對被稀釋），
# 低於 2% 門檻而完全沒被扣掉，整份濃度分母因此與手動分帳對不起來。
# 只用「完整名稱相等」比對，不用子字串 —— R5 的 'Hexane, 3-ethyl-4-methyl-'
# (C9，1.81%) 是產物不是溶劑，用子字串比對會被一起吃掉。
_SOLVENT_EXACT_NAMES = frozenset((
    'HEXANE', 'N-HEXANE',
    'PENTANE, 3-METHYL-', '3-METHYLPENTANE', 'PENTANE, 2-METHYL-', '2-METHYLPENTANE',
    'CYCLOPENTANE, METHYL-', 'METHYLCYCLOPENTANE',
))


def _is_solvent_name(name) -> bool:
    """名稱是否正好是 n-hexane 溶劑（或圖庫對過載己烷峰常回報的同分異構物）。"""
    n = str(name).upper().strip().rstrip('#*. ').strip()
    return n in _SOLVENT_EXACT_NAMES



# ── PubChem Hybrid 查詢（僅對 others 觸發）────────────────────────────
_pubchem_cache: dict = {}   # in-memory cache，每次重啟清空
PUBCHEM_CACHE_MAX = 5000    # 上限，避免長時間執行後無限成長吃光記憶體


def _cache_put(key: str, value: str) -> None:
    """寫入查詢快取；超過上限就整批清掉重來（用量小，不值得為此引入 LRU）。"""
    if len(_pubchem_cache) >= PUBCHEM_CACHE_MAX:
        _pubchem_cache.clear()
    _pubchem_cache[key] = value

PUBCHEM_HTTP_TIMEOUT   = 4     # 單次查詢逾時（秒）
PUBCHEM_MAX_WORKERS    = 8     # 並行查詢數（純網路等待，用執行緒即可）
PUBCHEM_RETRIES        = 2     # 暫時性失敗（429 限流 / 連線中斷）重試次數
# 每次上傳的 PubChem 查詢總時間上限（秒）。
# 必須明顯小於 gunicorn 的 --timeout（Render 預設只有 30 秒），
# 讓後面的寫檔（wb.save 是最慢的一步）與回傳都還有餘裕。
#
# 關鍵字規則涵蓋率提高後，實務上幾乎不會有化合物走到這裡
# （R1／R2 實測為 0 筆），所以這個預算是「保險絲」而不是常態路徑。
PUBCHEM_TIME_BUDGET    = 20


class _PubChemBudget:
    """
    整份活頁簿共用的 PubChem 查詢時間預算。

    遠端（Render + gunicorn）對單一 request 有 worker timeout；若某份檔案含大量
    無法用關鍵字判定的化合物，逐一查 PubChem 會拖垮整個 request，最後回傳的是
    平台的 HTML 錯誤頁而不是 JSON。這裡設一個硬上限：預算用完後不再發出網路查詢，
    剩下的一律留在 'others'，並回報略過筆數，讓使用者知道有部分未做二次判定。
    """

    def __init__(self, seconds: float = PUBCHEM_TIME_BUDGET):
        self._deadline = time.monotonic() + seconds
        self.skipped = 0

    def allows_lookup(self) -> bool:
        if time.monotonic() < self._deadline:
            return True
        self.skipped += 1
        return False


def _pubchem_lookup(name: str) -> str:
    """
    查 PubChem PUG REST API 取得分子式，依分子式二次分類。
    失敗時靜默回傳 'others'。
    """
    import urllib.request, urllib.parse, json as _json

    key = name.upper().strip()
    if key in _pubchem_cache:
        return _pubchem_cache[key]

    encoded = urllib.parse.quote(key)
    url = (f'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/'
           f'{encoded}/property/MolecularFormula,IUPACName/JSON')

    result = None
    for attempt in range(PUBCHEM_RETRIES + 1):
        try:
            with urllib.request.urlopen(url, timeout=PUBCHEM_HTTP_TIMEOUT) as r:
                data = _json.loads(r.read())
            props   = data['PropertyTable']['Properties'][0]
            formula = props.get('MolecularFormula', '')
            iupac   = (props.get('IUPACName') or '').upper()
            result  = _classify_by_formula(formula, iupac, key)
            break
        except Exception as e:
            # 404（PubChem 查無此化合物）不必重試，直接放棄；
            # 限流／連線問題才退避後再試。
            code = getattr(e, 'code', None)
            if code == 404 or attempt == PUBCHEM_RETRIES:
                break
            time.sleep(0.4 * (attempt + 1))

    if result is None:
        result = 'others'

    _cache_put(key, result)
    return result


def _classify_by_formula(formula: str, iupac: str, original: str) -> str:
    """
    依分子式 CₙHₘ 推斷化合物類別。
    規則：
      - 含苯環特徵（H 數明顯不飽和且 IUPAC 含 benzene/phenyl/naphthalene）→ Aromatic/Polyaromatic
      - CₙH₂ₙ₋₆ (苯環) 或 IUPAC 含 benzene → Aromatic
      - IUPAC 含 naphthalene/pyrene/fluorene/anthracene → Polyaromatic
      - CₙH₂ₙ₊₂ 且含 cyclo → cyclic  (實際上 cycloalkane = CₙH₂ₙ，見下)
      - CₙH₂ₙ    → cyclic（若 iupac 含 cycl）或 Alkene
      - CₙH₂ₙ₊₂  → alkane；有 methyl branch → iso-alkane，否則 → n-alkane
    """
    import re

    # PubChem 回傳的 IUPAC 名稱本身也是化學名，先用同一套關鍵字規則判一次。
    # 這樣兩條路徑（原始名稱 / PubChem 名稱）永遠遵循同一套慣例，不會各說各話。
    # 實測顯示，先前兩邊各自維護關鍵字，正是大多數判定不一致的來源
    # （例如 acenaphthylene 在這裡漏列、biphenyl 被 IUPAC 寫成 phenylbenzene 而誤判單環）。
    if iupac:
        cat, confident = _classify_keyword(iupac)
        if confident:
            return cat

    # 只含 C/H（可含 O/N）才適用下面的分子式啟發式；含鹵素、硫等一律 others
    m = re.match(r'^C(\d+)H(\d+)((?:O\d*)?(?:N\d*)?)$', formula)
    if not m:
        return 'others'

    c = int(m.group(1))
    h = int(m.group(2))
    hetero = m.group(3)   # 含 O/N → 雜原子化合物（脂肪醇、醛、酮、酯、胺）

    # 帶 O/N 的非芳香化合物 → 氧化物 / 含雜原子，非純碳氫類 → others
    # （修正：脂肪醇 CₙH₂ₙ₊₂O 會被誤判成 n-alkane，因為過去只看 C/H 忽略了 O）
    if hetero:
        return 'others'

    # 單環芳烴（苯環 CₙH₂ₙ₋₆）。這只是最後的啟發式：CₙH₂ₙ₋₆ 不保證是苯環
    # （fulvene 類的環狀多烯也符合），所以放在關鍵字判定之後才用。
    if h == 2 * c - 6:
        return 'Aromatic'

    # Cycloalkane CₙH₂ₙ + cycl in name
    if h == 2 * c and 'CYCL' in iupac:
        return 'cyclic'

    # Alkene CₙH₂ₙ（不含 cyclo）
    if h == 2 * c:
        return 'Alkene'

    # Alkane CₙH₂ₙ₊₂
    if h == 2 * c + 2:
        # 與 _classify_keyword 第 7 條用同一套判準（見該處說明），
        # 避免兩條路徑對 "3,4-dimethyloctane" 這種多重取代基各說各話。
        if any(x in iupac for x in _ALKYL_SUBSTITUENTS):
            return 'iso-alkane'
        return 'n-alkane'

    return 'others'

# IUPAC 碳數字根 C1–C50。原本的清單只到 OCTADECANE(C18)，導致 GC-MS 常見的
# 長鏈化合物（HENEICOSANE C21、DOTRIACONTANE C32、TRITETRACONTANE C43…）
# 全部落到 others、被迫去查 PubChem。字根 + 'ANE'/'ENE' 即可涵蓋烷與烯。
_CHAIN_STEMS = [
    'METH','ETH','PROP','BUT','PENT','HEX','HEPT','OCT','NON','DEC',
    'UNDEC','DODEC','TRIDEC','TETRADEC','PENTADEC','HEXADEC','HEPTADEC',
    'OCTADEC','NONADEC',
    'ICOS','EICOS','HENICOS','HENEICOS','DOCOS','TRICOS','TETRACOS',
    'PENTACOS','HEXACOS','HEPTACOS','OCTACOS','NONACOS',
    'TRIACONT','HENTRIACONT','DOTRIACONT','TRITRIACONT','TETRATRIACONT',
    'PENTATRIACONT','HEXATRIACONT','HEPTATRIACONT','OCTATRIACONT','NONATRIACONT',
    'TETRACONT','HENTETRACONT','DOTETRACONT','TRITETRACONT','TETRATETRACONT',
    'PENTATETRACONT','HEXATETRACONT','HEPTATETRACONT','OCTATETRACONT','NONATETRACONT',
    'PENTACONT',
]
_ALKANE_NAMES = tuple(s + 'ANE' for s in _CHAIN_STEMS)
_ALKENE_NAMES = tuple(s + 'ENE' for s in _CHAIN_STEMS)

# 烷基取代基字根（METHYL、ETHYL、…、OCTYL、DODECYL…）。
# 在「已判定為烷類」的名稱裡出現任何一個，就代表主鏈上掛了支鏈 → iso-alkane。
#
# 舊版用 r'\d\s*-\s*(METHYL|ETHYL|PROPYL|BUTYL)' 比對，只吃得到「單一取代基」
# （"Octane, 3-methyl-" ✓）。碰到多重取代基時，數字後面接的是 DI/TRI 而不是
# METHYL，整批漏判成 n-alkane：
#     OCTANE, 3,4-DIMETHYL-        → n-alkane（錯）
#     DODECANE, 2,6,10-TRIMETHYL-  → n-alkane（錯）
#     DODECANE, 4,9-DIPROPYL-      → n-alkane（錯）
#     HEPTADECANE, 9-OCTYL-        → n-alkane（錯，OCTYL 根本不在舊清單裡）
# 這是 R5 實測 i/n 被算成 0.17（正解 0.90）的唯一主因，共 14.05% 濃度跑錯類別。
_ALKYL_SUBSTITUENTS = tuple(s + 'YL' for s in _CHAIN_STEMS)

# 多環芳烴字根，含 NIST/CAS 圖庫實際出現過的拼字變體。
# R5 的 "METHYL-ETHYL-NAPHTALENE"（NAPHTHALENE 少一個 H）在舊版比對不到
# 'NAPHTHALEN'，一路掉到 others 再去查 PubChem（也查不到）→ 0.26% 歸錯類。
_NAPHTHALENE_TOKENS = ('NAPHTHALEN', 'NAPHTALEN', 'NAPTHALEN', 'NAPHTHALIN')
_PAH_TOKENS = _NAPHTHALENE_TOKENS + (
    'ANTHRACEN', 'ANTRACEN', 'PHENANTHREN', 'PHENANTREN', 'PYREN',
    'FLUOREN', 'BIPHENYL', 'DIPHENYL', 'TRIPHENYL', 'ACENAPHTH',
    'AZULEN', 'FLUORANTHEN', 'FLUORANTEN', 'PERYLEN', 'CHRYSEN', 'DIBENZO',
)

# 非碳氫官能基（含 O / N / Si / S）。在烴類分類體系中一律歸 others。
# 這條規則排在芳香環之前，所以帶環的含氧物（phenanthrenol、phenyl ketone、
# anthracenyl-propanol…）也會在這裡被攔下 → others，與 phenol／naphthol
# 的處理一致（PIONA 的 aromatic 只收芳香「烴」）。
_HETERO_TOKENS = (
    'ACID', 'ESTER', 'ETHER', 'AMINE', 'AMINO', 'AMIDE', 'NITRIL', 'NITRO',
    'SILANE', 'SILOXANE', 'SILOXY', 'THIO', 'SULFON', 'SULFID', 'SULFIN',
    'SULFANYL', 'MERCAPTO', 'PHOSPH', 'CYANO',
    # 鹵素：含 Cl/Br/F/I 的化合物同樣不是純碳氫。
    # （'FLUORO' 不會誤中 'FLUORENE'，且芳香環判定本來就排在前面）
    'CHLORO', 'BROMO', 'FLUORO', 'IODO', 'CHLORIDE', 'BROMIDE', 'IODIDE',
    'FURAN', 'PYRAN', 'OXEPINE', 'OXIRANE', 'TRIAZIN', 'PYRIDIN', 'PYRROL',
    'IMIDAZOL', 'THIOPHEN', 'ANOL', 'ANAL', 'ANONE', 'ANOATE', 'OATE',
    'ALDEHYDE', 'KETONE', 'OXO', 'HYDROXY', 'CARBOXY',
    # 醚／過氧／環氧一律用字根 'OXY' / 'OXA' 一次涵蓋（取代原本逐一列舉的
    # METHOXY / ETHOXY —— 那份清單漏掉 ETHENYLOXY、PHENOXY、EPOXY、PEROXY…）。
    # R10 實測：'Octadecane, 1-(ethenyloxy)-' 因為 'ETHENYLOXY' 比對不到
    # 'ETHOXY'，一路掉到第 7 條被判成 n-alkane（正解 others）。
    # 'OXA' 吃 oxa- 雜環（'1-(6-OXABICYCLO[3.1.0]HEX-1-YL)ETHANONE'）。
    'OXY', 'OXA',
    # 酚類與酯類。CAS 倒裝寫法（"PHENOL, 2-METHYL-"）不會以 OL 結尾，
    # 只靠 n.endswith('OL') 會漏；phthalate 是 GC-MS 最常見的塑化劑污染，
    # 它是酯不是芳香烴，計進 aromatic 會直接灌水 aromatic 佔比。
    'PHENOL', 'CRESOL', 'NAPHTHOL', 'PHTHALATE', 'FORMATE', 'ACETATE',
    # 內酯（-OLIDE）、脂肪酸甘油酯（NIST 對長鏈酯常回報成 "…SAFFLOWER OIL"）
    'OLIDE', 'OLEIC', 'OLEATE', 'STEARAT', 'PALMITAT',
)

# 醇的字尾。同樣不能只比對 'OL'：CAS 倒裝寫法把羥基寫在名稱中段，
# 現有的 n.endswith('OL') 完全吃不到：
#     BICYCLO[3.1.1]HEPTAN-3-OL, 6,6-DIMETHYL-2-METHYLENE-, [1S-…]-  → -3-OL ✓
#     3-NONEN-1-OL, (E)-                                             → -1-OL ✓
#     2-PROPYN-1-OL, 3-P-TOLYL-                                      → -1-OL ✓
_ALCOHOL_RE = re.compile(r'(?:-\d[\d,]*-OL|ANOL|ENOL|INOL|YNOL)(?![A-Z])')

# 酮／醛的字尾。不能只比對子字串 'ONE' —— 'NONENE'（壬烯）本身就含 'ONE'，
# 會把一整批烯類洗成 others。只認「數字-ONE」「…ANONE」「…ENONE」三種寫法：
#     DECA-5,7-DIEN-1-ONE      → -1-ONE      ✓
#     2-TETRADECANONE          → ANONE       ✓
#     BICYCLO[3.1.1]HEPTAN-3-ONE → -3-ONE    ✓
#     CIS-3-NONENE             → 都不中      ✓（正確，這是烯）
_KETONE_RE = re.compile(r'(?:-\d[\d,]*-ONE|ANONE|ENONE|INONE)(?![A-Z])')

# 環狀多烯的不飽和字尾（bicyclo/tricyclo/spiro 系統常見）。
# 不含字尾 E，才能同時吃到 "…-TRIENE" 與德文式寫法 "…-TRIEN,"。
_POLYENE_TOKENS = ('DIEN', 'TRIEN', 'TETRAEN', 'PENTAEN', 'HEXAEN', 'HEPTAEN')

# 不飽和取代基字根（掛在主鏈或環上的 C=C）。VINYL/ETHENYL/ALLYL 之外，
# 'ENYL' 一次涵蓋 PROPENYL、HEXENYL、OCTENYL… 等所有烯基。
# 注意 'ENYL' 也會中 'PHENYL'，所以這組字根只能在芳香環判定「之後」使用。
_UNSAT_SUBSTITUENTS = ('VINYL', 'ETHENYL', 'ALLYL', 'ENYL', 'YLIDENE', 'METHYLENE')

# 六員環上的 1,3,5-三烯 = 苯環（1,3,5-cyclohexatriene 就是苯的舊式命名）。
# NIST 有時會用 von Baeyer 系統把苯環藏在 bicyclo 名稱裡：
#     7-BUTYLBICYCLO[4.2.0]OCTA-1,3,5-TRIENE = 丁基苯并環丁烯（芳香）
#     BICYCLO[4.3.0]NONA-1,3,5-TRIENE        = indane 骨架（芳香）
# 但 1,3,5-CYCLOHEPTATRIENE（七員環）不是芳香烴，所以只認「六員環」的寫法：
# bicyclo[4.x.0] 的第一個橋長 4 ＋ 融合處 2 個碳 = 六員環。
_ARENE_RE = re.compile(
    r'(?:BICYCLO\[4\.\d+\.0\][A-Z]*-?1,3,5-TRIEN|CYCLOHEXA-?1,3,5-TRIEN'
    r'|1,3,5-CYCLOHEXATRIEN)')


def _is_unsaturated(n: str) -> bool:
    """名稱是否帶非芳香的 C=C 雙鍵。呼叫端必須先排除芳香烴（'ENYL' 會中 PHENYL）。"""
    return ('-ENE' in n
            or any(x in n for x in _POLYENE_TOKENS)
            or any(x in n for x in _ALKENE_NAMES)
            or any(x in n for x in _UNSAT_SUBSTITUENTS))


def _classify_keyword(name: str):
    """
    規則式分類，回傳 (category, confident)。

    confident=False 只代表「所有規則都沒對上、真的不知道」，這種才值得去查 PubChem。
    像 furan / silane / 醇 / 酸 / 酯 這類含雜原子的化合物，others 本身就是正確答案
    （烴類分類體系裡沒有對應類別），屬於 confident=True，不該浪費一次網路查詢。

    判定順序（先到先贏，順序本身就是分類慣例，改動順序＝改動分類標準）：
      1. 含氧／雜原子官能基（O / N / S / Si / 鹵素）→ others
      2. 部分氫化的萘／茚（只剩 1 個芳香環）        → Aromatic
      3. 多環芳烴（≥2 個芳香環）                    → Polyaromatic
      4. 單環芳烴                                   → Aromatic
      5. 不飽和（C=C，含環烯）                      → Alkene
      6. 飽和環烷（只有單鍵才算）                    → cyclic
      7. 烷                                         → n-alkane / iso-alkane

    ⚠ 第 1 條被提到最前面（舊版排在芳香環判定之後）。烴類分類體系（PIONA）
      的 aromatic 指的是「芳香『烴』」，含氧的芳香化合物屬於含氧物：
          .BETA.-PHENETHYL PHENYL ETHER          舊 Aromatic → 新 others
          1-Butanone, 2-methyl-1-phenyl-         舊 Aromatic → 新 others
          (1-METHOXYMETHOXY-BUT-2-ENYL)-BENZENE  舊 Aromatic → 新 others
      這與實驗室手動分類的慣例一致（含 O/鹵素一律 others）。

    ⚠ 2026-08-26：不飽和（第 5 條）被提到環烷（第 6 條）之前。
      cyclic 的定義是「只有飽和單鍵的環」，環烯屬於不飽和 → Alkene：
          CYCLOPENTENE, 4,4-DIMETHYL-            cyclic → Alkene
          Cyclohexene, 1-methyl-                 cyclic → Alkene
          7-BUTYL-1,3,5-CYCLOHEPTATRIENE         cyclic → Alkene
          6,6-DIMETHYL-2-VINYL-BICYCLO[3.1.1]HEPT-2-ENE  cyclic → Alkene
      舊順序只是因為「CYCLOPENTENE 不含字串 PENTANE」才碰巧判對，
      碰到 "CYCLOHEXANE, ETHENYL-"（乙烯基環己烷）這種寫法就會判錯。
    """
    n = name.upper().strip()

    # ── 1. 含氧／雜原子官能基（醇/醛/酮/酸/酯/醚/胺/矽/鹵素/含氧氮雜環）→ others ──
    #    這些不是碳氫化合物，在烴類分類體系中沒有對應類別，others 就是正解，
    #    不必再查 PubChem（標記 confident=True）。
    if (any(x in n for x in _HETERO_TOKENS)
            or n.endswith('OL')
            or _ALCOHOL_RE.search(n)
            or _KETONE_RE.search(n)):
        return 'others', True

    # ── 2. 部分氫化的萘/茚 = 只剩「一個」芳香環 → 單環 Aromatic ──────────
    #    tetralin(四氫萘)、dihydronaphthalene、indane、indene、indacene。
    #    採「芳香環數目」慣例:此類只有 1 個芳香環,不列 Polyaromatic(PAH 需 ≥2 芳香環)。
    #    相容 CAS 倒裝 "NAPHTHALENE, 1,2,3,4-TETRAHYDRO-" 與正寫 "...TETRAHYDRONAPHTHALENE"。
    #    一律用「字根」比對(NAPHTHALEN 而非 NAPHTHALENE)，否則衍生物寫法
    #    (NAPHTHALENYL、INDEN-1-ONE…)會整批漏掉。
    _is_naph = any(x in n for x in _NAPHTHALENE_TOKENS)
    if _is_naph and ('DECAHYDRO' in n or 'DECALIN' in n):
        return 'cyclic', True                 # 全飽和 decalin → 環烷
    if ((_is_naph and ('TETRAHYDRO' in n or 'DIHYDRO' in n))
            or 'TETRALIN' in n
            or any(x in n for x in ('INDEN', 'INDAN', 'INDACENE'))):
        return 'Aromatic', True

    # ── 3. 真正多環芳烴(≥2 個芳香環)→ Polyaromatic ────────────────────
    if any(x in n for x in _PAH_TOKENS):
        return 'Polyaromatic', True
    #    數苯環：PubChem 的 IUPAC 會把 biphenyl 拆寫成
    #    "1-METHYL-2-(3-METHYLPHENYL)BENZENE"（中間夾括號，比對不到 BIPHENYL），
    #    但只要出現兩個苯環標記就是二芳香環系統。
    if n.count('PHENYL') >= 2 or ('PHENYL' in n and 'BENZENE' in n):
        return 'Polyaromatic', True

    # ── 4. 單環芳烴(苯環)→ Aromatic ──────────────────────────────────
    #    ANNULENE = 大環芳香烴(Hückel 4n+2)。
    #    苯環＋環烷的 naphtheno-aromatic（CYCLOPENTANE, 1,1-DIMETHYL-3-PHENYL-）
    #    依 PIONA 慣例算 aromatic，所以這條刻意排在環烷之前。
    if any(x in n for x in ('BENZENE','BENZO','BENZYL','TOLUEN','XYLEN','STYREN',
                              'PHENYL','CUMENE','CYMENE','ANNULENE')):
        return 'Aromatic', True
    #    苯環被 von Baeyer 系統命名藏起來的寫法（六員環上的 1,3,5-三烯）。
    #    必須排在第 5 條之前，否則 '…OCTA-1,3,5-TRIENE' 會被當成環狀多烯判成 Alkene。
    if _ARENE_RE.search(n):
        return 'Aromatic', True

    # ── 5. 不飽和（C=C）→ Alkene ─────────────────────────────────────
    #    含 bicyclo/tricyclo/spiro 的環狀多烯，以及環烯（cyclohexene…）。
    #    這條排在環烷之前：cyclic 依定義只收「全飽和單鍵」的環，
    #    環上只要有一個 C=C 就屬於不飽和，歸 Alkene。
    if _is_unsaturated(n):
        return 'Alkene', True

    # ── 6. 飽和環烷 → cyclic ─────────────────────────────────────────
    #    舊版只列舉 CYCLOPROPANE~CYCLOOCTANE 六個環，中大環全部漏掉：
    #    CYCLODODECANE 因為字串含 'DODECANE'，會被第 7 條判成 n-alkane。
    #    改成「含 CYCLO 且含任一烷類字根」，一次涵蓋 C3~C50 環與 bicyclo/spiro 烷。
    if 'CYCLO' in n and any(x in n for x in _ALKANE_NAMES):
        return 'cyclic', True
    if 'DECAHYDRO' in n:
        return 'cyclic', True

    # ── 7. 烷類 → n-alkane / iso-alkane ──────────────────────────────
    #    「名稱裡出現任何烷基取代基字根 → 有支鏈 → iso」。
    #    這比舊版的 locant 正則穩健：不必去猜取代基前面接的是 "3-"、"3,4-DI"
    #    還是 "2,6,10-TRI"，也不會因為取代基不是甲基（OCTYL、PROPYL）就漏判。
    if any(x in n for x in _ALKANE_NAMES):
        if (any(x in n for x in _ALKYL_SUBSTITUENTS)
                or 'ISO' in n or 'NEO' in n):
            return 'iso-alkane', True
        return 'n-alkane', True

    # 所有規則都沒對上 → 交給 PubChem 用分子式二次判定
    return 'others', False


def auto_classify(name: str) -> str:
    """規則式分類的對外介面（只回傳類別）。"""
    return _classify_keyword(name)[0]


def _prefetch_pubchem(names, budget=None) -> None:
    """
    把「關鍵字判不出來」的化合物名稱一次並行查完，結果寫進 _pubchem_cache。

    這些查詢純粹是等網路回應，逐一串接跑會直接決定整個 request 的長度
    （實測 80 個化合物循序查要 ~20 秒，是逾時被 SIGKILL 的主因）。
    改用執行緒池並行，同樣的量降到數秒。
    """
    todo = []
    seen = set()
    for name in names:
        key = str(name).upper().strip()
        if key and key not in seen and key not in _pubchem_cache:
            seen.add(key)
            todo.append(name)
    if not todo:
        return

    from concurrent.futures import ThreadPoolExecutor

    def _one(nm):
        if budget is not None and not budget.allows_lookup():
            return
        _pubchem_lookup(nm)          # 內部會自行寫入快取

    workers = min(PUBCHEM_MAX_WORKERS, len(todo))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        list(pool.map(_one, todo))


def _pubchem_fallback(name: str, budget=None) -> str:
    """
    先跑規則式分類；只有「規則完全判不出來」(confident=False) 才查 PubChem。
    已快取的名稱直接取用（不佔預算）；未快取且預算已用罄時，保持 'others'。
    """
    result, confident = _classify_keyword(name)
    if confident:
        return result

    key = name.upper().strip()
    if key in _pubchem_cache:
        return _pubchem_cache[key]

    if budget is not None and not budget.allows_lookup():
        return 'others'

    return _pubchem_lookup(name)


# ── 溶劑峰偵測 ────────────────────────────────────────────────────────
def _resolve_solvent_peaks(data, opts):
    """
    決定哪幾列是溶劑峰，回傳 index 清單。

    mode='manual'：使用者直接指定 Peak#（"1,2" 或 "1-3"），完全照做。
    mode='auto'  ：從層析圖「最前端」連續往後掃，(a) 成立且 (b)(c) 任一成立
                   就標成溶劑，遇到第一個不符合的峰立刻停手：
                     (a) RT 早於 C8 邊界   —— 溶劑一定跑在最輕的產物之前
                     (b) 單峰 Conc% ≥ 門檻 —— 溶劑是過載峰，濃度遠高於產物
                     (c) 名稱正好是己烷或其 C6H14 同分異構物 —— 見 _is_solvent_name
                   「連續」這個條件很重要：溶劑主峰＋拖尾是相鄰的一段，
                   一旦回到正常大小的峰就代表拖尾結束，後面都是產物。
    mode='off'   ：不扣除（保留舊行為）。

    R5 實測：峰1 (RT 2.016, 3.72%)、峰2 (RT 2.070, 5.76%) 判為溶劑；
    峰3 (RT 2.143, 1.81%, 'Hexane, 3-ethyl-4-methyl-') 濃度低於門檻、
    名稱也不是純己烷，掃描在此停止 —— 與手動判讀結果一致。
    R10 實測：峰1 (RT 2.016, 0.91%, 'PENTANE, 3-METHYL-') 靠條件 (c) 抓到；
    峰2 ('1-HEXENE, 2-METHYL-') 兩個條件都不成立而停止 —— 與 Jack 手動
    "without Hexane" 分帳（只扣 0.91%）完全一致。
    """
    mode = (opts.get('solvent_mode') or 'auto').lower()
    if mode == 'off':
        return []

    if mode == 'manual':
        wanted = _parse_peak_list(opts.get('solvent_peaks') or '')
        if not wanted:
            return []
        col = 'Peak#' if 'Peak#' in data.columns else None
        idx = []
        for i in data.index:
            pk = data.at[i, col] if col else (i + 1)
            try:
                pk = int(float(pk))
            except (TypeError, ValueError):
                continue
            if pk in wanted:
                idx.append(i)
        return idx

    rt_max   = opts.get('rt_c8_min', RT_C8_MINUS_MAX)
    min_conc = opts.get('solvent_min_conc', SOLVENT_MIN_CONC)
    idx = []
    for i in data.index:
        rt = data.at[i, 'Ret.Time']
        cc = data.at[i, 'Conc.']
        if pd.isna(rt) or rt >= rt_max:
            break
        big    = pd.notna(cc) and cc >= min_conc
        hexane = _is_solvent_name(data.at[i, 'Name'])
        if not (big or hexane):
            break
        idx.append(i)
    return idx


def _parse_peak_list(text):
    """把 "1,2,5-7" 解析成 {1,2,5,6,7}；解析不了的片段直接忽略。"""
    out = set()
    for part in str(text).replace('，', ',').split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a, _, b = part.partition('-')
            try:
                out.update(range(int(a), int(b) + 1))
            except ValueError:
                continue
        else:
            try:
                out.add(int(part))
            except ValueError:
                continue
    return out


# ── 名稱碳數 vs 滯留時間一致性稽核 ────────────────────────────────────
_STEM_CARBONS = {
    'METH':1,'ETH':2,'PROP':3,'BUT':4,'PENT':5,'HEX':6,'HEPT':7,'OCT':8,'NON':9,
    'DEC':10,'UNDEC':11,'DODEC':12,'TRIDEC':13,'TETRADEC':14,'PENTADEC':15,
    'HEXADEC':16,'HEPTADEC':17,'OCTADEC':18,'NONADEC':19,'ICOS':20,'EICOS':20,
    'HENICOS':21,'HENEICOS':21,'DOCOS':22,'TRICOS':23,'TETRACOS':24,'PENTACOS':25,
    'HEXACOS':26,'HEPTACOS':27,'OCTACOS':28,'NONACOS':29,'TRIACONT':30,
    'HENTRIACONT':31,'DOTRIACONT':32,'TRITRIACONT':33,'TETRATRIACONT':34,
    'PENTATRIACONT':35,'HEXATRIACONT':36,'HEPTATRIACONT':37,'OCTATRIACONT':38,
    'NONATRIACONT':39,'TETRACONT':40,'HENTETRACONT':41,'DOTETRACONT':42,
    'TRITETRACONT':43,'TETRATETRACONT':44,'PENTATETRACONT':45,'HEXATETRACONT':46,
    'HEPTATETRACONT':47,'OCTATETRACONT':48,'NONATETRACONT':49,'PENTACONT':50,
}
_RING_CARBONS = {
    'NAPHTHALEN':10,'NAPHTALEN':10,'NAPTHALEN':10,'ANTHRACEN':14,'PHENANTHREN':14,
    'FLUOREN':13,'PYREN':16,'BIPHENYL':12,'INDEN':9,'INDAN':9,'TETRALIN':10,
    'AZULEN':10,'TOLUEN':7,'XYLEN':8,'STYREN':8,'CUMENE':9,'CYMENE':10,'BENZENE':6,
}
_SUBST_CARBONS = {
    'METHYL':1,'ETHYL':2,'PROPYL':3,'BUTYL':4,'PENTYL':5,'HEXYL':6,'HEPTYL':7,
    'OCTYL':8,'NONYL':9,'DECYL':10,'UNDECYL':11,'DODECYL':12,'TRIDECYL':13,
    'TETRADECYL':14,'PENTADECYL':15,'HEXADECYL':16,'PHENYL':6,'BENZYL':7,'TOLYL':7,
    # 不飽和取代基。注意 'PROPENYL' 不會被 'PROPYL' 吃掉（中間多一個 E），
    # 所以兩份清單可以並存而不會重複計數。
    'VINYL':2,'ETHENYL':2,'ALLYL':3,'PROPENYL':3,'BUTENYL':4,'PENTENYL':5,
    'HEXENYL':6,'HEPTENYL':7,'OCTENYL':8,
}

# 只有脂肪族（烷/烯/環烷）才適用「RT ↔ 碳數」線性關係——這條線是拿
# n-alkane 校正出來的。芳香烴在非極性管柱上的滯留指數系統性偏高
# （naphthalene C10 落在 n-C11 與 n-C12 之間），拿同一條線去比，
# 會把一整批正常的萘系化合物誤報成「圖庫誤判」。
_CARBON_CHECK_CATEGORIES = ('n-alkane', 'iso-alkane', 'Alkene', 'cyclic')
_MULTIPLIERS = {'DI':2,'TRI':3,'TETRA':4,'PENTA':5,'HEXA':6,'HEPTA':7}

# 名稱碳數與 RT 推得的碳數相差多少才示警。設 3 是為了讓正常的
# 支鏈／異構物落差（實測 ≤2）不會洗版，只抓真正離譜的圖庫誤判。
CARBON_MISMATCH_TOL = 3


def _name_carbon_count(name):
    """
    從化合物名稱粗估碳數：主鏈（或環系統）碳數 + 所有取代基碳數。
    純啟發式，只用來做「差太多就示警」的稽核，不參與分類。
    判不出來回傳 None。
    """
    import re as _re
    n = str(name).upper()

    parent = 0
    for stem, c in _STEM_CARBONS.items():
        for suffix in ('ANE', 'ENE', 'YNE'):
            if stem + suffix in n:
                parent = max(parent, c)
    for ring, c in _RING_CARBONS.items():
        if ring in n:
            parent = max(parent, c)
    if not parent:
        return None

    subs = 0
    alt = '|'.join(sorted(_SUBST_CARBONS, key=len, reverse=True))
    mul = '|'.join(sorted(_MULTIPLIERS, key=len, reverse=True))
    for m in _re.finditer(r'(%s)?(%s)' % (mul, alt), n):
        subs += _MULTIPLIERS.get(m.group(1), 1) * _SUBST_CARBONS[m.group(2)]
    return parent + subs


def _carbon_from_rt(rt, rt_c8, rt_c16):
    """
    用 C8 與 C16 兩個 n-alkane 標記峰的滯留時間線性內插推碳數。
    程式升溫下 n-alkane 的 RT 對碳數接近線性（R5 實測誤差 <0.5 個碳）。
    """
    if rt_c16 <= rt_c8:
        return None
    return 8.0 + (rt - rt_c8) / (rt_c16 - rt_c8) * 8.0


# ── n-alkane 滯留時間標尺 ──────────────────────────────────────────────
# 直鏈烷（n-alkane）在同一套 GC 方法下，每個碳數只會出現在「一個」固定的
# 滯留時間。NIST 圖庫對長鏈飽和烷的質譜幾乎分不出直鏈與支鏈（碎片圖譜近乎
# 相同），常常把一根 iso-alkane 回報成 "DODECANE"、"Pentadecane" 這種直鏈名。
# 只信名稱 → n-alkane 被灌水、i/n 被低估。
#
# 判準：名稱是「沒有取代基的直鏈烷」時，才去標尺上找它的碳數對應的 RT；
#       同一個碳數若有好幾根候選峰，離標尺最近的那一根才是 n-alkane，
#       其餘一律降級成 iso-alkane（同碳數的其他飽和異構物）。
#
# 預設值來自實驗室的 n-alkane 標準品實測表（C8–C18，Octane → Octadecane）。
# 這條標尺跟著管柱與溫度程式走，換方法就要換數字 —— 所以做成可在前端覆寫。
DEFAULT_NALKANE_RT = {
    8: 2.2, 9: 3.6, 10: 6.4, 11: 8.9, 12: 10.9, 13: 12.5,
    14: 14.2, 15: 16.0, 16: 17.6, 17: 19.2, 18: 20.7,
}

# 候選峰離標尺多遠就完全不採信（分鐘）。設 1.5 是因為標尺上相鄰碳數
# 最窄的間隔約 1.4–1.6 分鐘，超過一個碳的距離就不可能是同一根峰。
# R10 實測最大偏移是 C8 的 0.53 分鐘（該次分析前段整體偏晚），落在容差內。
NALKANE_RT_CAP = 1.5

# 標定成功、但偏移大到值得人工複核的門檻（分鐘）。
# R10 實測：C8~C16 的偏移在 -0.05 ~ +0.53 之間，只有 C18 是 -1.05
# —— 而同一份資料在 20.511 分（正好是 C18 標尺位置）有一根被圖庫叫成
# "NONADECANE" 的峰。這正是 NIST 分不出 C18/C19 直鏈的典型狀況，
# 程式不自作主張改名，但一定要把這種偏移標出來給人看。
NALKANE_RT_WARN = 0.6

_NALKANE_STEM_BY_CARBON = {}
for _stem, _c in _STEM_CARBONS.items():
    _NALKANE_STEM_BY_CARBON.setdefault(_c, _stem)


def _plain_nalkane_carbons(name):
    """
    名稱是否為「沒有任何取代基的直鏈烷」；是就回傳碳數，否則 None。

        'OCTANE'                     → 8
        'n-Hexadecane'               → 16
        'Heneicosane'                → 21
        'Pentadecane '               → 15
        'DODECANE, 4-METHYL-'        → None（有取代基）
        'CYCLODODECANE'              → None（環烷）
    比對前先去掉 NIST 的品質標記（結尾的 '#'、'*'）與 'n-' 前綴。
    """
    n = str(name).upper().strip().rstrip('#*. ').strip()
    if n.startswith('N-'):
        n = n[2:]
    if not n.isalpha():
        return None
    for stem, c in _STEM_CARBONS.items():
        if n == stem + 'ANE':
            return c
    return None


def _parse_rt_ruler(text):
    """
    解析前端傳來的標尺文字，回傳 {碳數: RT}。接受這些寫法（可混用）：
        "C8 2.2, C9 3.6"      "8:2.2 9:3.6"      每行一組 "C10=6.4"
    解析不出來的片段直接忽略；整份都解析不出來就回傳 None（呼叫端用預設值）。
    """
    if not text or not str(text).strip():
        return None
    out = {}
    for m in re.finditer(r'C?\s*(\d{1,2})\s*[:=\s]\s*(\d+(?:\.\d+)?)',
                         str(text).upper().replace('，', ',')):
        c, rt = int(m.group(1)), float(m.group(2))
        if 1 <= c <= 60 and rt > 0:
            out[c] = rt
    return out or None


def _plan_nalkane_rt_ruler(items, ruler, cap=NALKANE_RT_CAP, mode='strict'):
    """
    對「規則判成 n-alkane」的峰執行滯留時間校驗。

    items : list of (key, rt, name) —— 只傳目前類別為 n-alkane 的峰。
    ruler : {碳數: 標準 RT}
    cap   : 離標尺多遠就不採信（分鐘）
    mode  : 'strict' 標尺沒有的碳數一律降級 iso-alkane（預設，對應手動分類慣例）
            'range'  只在標尺涵蓋的碳數區間內校驗，區間外沿用名稱判定

    回傳 (demoted, anchors, unverified)
      demoted    : [{key, rt, name, carbon, ref_rt, delta, why, by_rt}] 要改判 iso-alkane 的峰
                   why 是代碼而非句子（'no_ruler' / 'too_far' / 'taken'），
                   因為同一份結果要同時餵給 Excel 報告（英文）與前端（中／英），
                   在這裡寫死句子等於把語言綁進資料。
      anchors    : [{carbon, ref_rt, key, rt, name, delta, warn}] 每個碳數對到的標記峰
      unverified : [{key, rt, name, reason}] 無法校驗、維持 n-alkane 的峰
    """
    by_carbon, unverified = {}, []
    for key, rt, name in items:
        c = _plain_nalkane_carbons(name)
        if c is None or rt is None:
            unverified.append({'key': key, 'rt': rt, 'name': str(name),
                               'why': 'unparsed'})
            continue
        by_carbon.setdefault(c, []).append((key, float(rt), str(name)))

    c_lo = min(ruler) if ruler else None
    c_hi = max(ruler) if ruler else None

    demoted, anchors = [], []
    for c in sorted(by_carbon):
        cands = by_carbon[c]
        ref = ruler.get(c)
        if ref is None:
            outside = (c_lo is None) or (c < c_lo) or (c > c_hi)
            if mode == 'range' and outside:
                # 標尺涵蓋範圍之外 —— 不校驗也不降級，只登記成「未校驗」。
                for key, rt, name in cands:
                    unverified.append({'key': key, 'rt': rt, 'name': name,
                                       'why': 'out_of_range', 'carbon': c,
                                       'c_lo': c_lo, 'c_hi': c_hi})
                continue
            # strict：標尺沒有這個碳數 → 無從確認是不是直鏈，保守降級成 iso-alkane。
            # 要讓某個碳數也能判 n-alkane，就把它的實測 RT 加進標尺。
            for key, rt, name in cands:
                demoted.append({'key': key, 'rt': rt, 'name': name, 'carbon': c,
                                'ref_rt': None, 'delta': None, 'by_rt': None,
                                'why': 'no_ruler'})
            continue
        cands.sort(key=lambda t: abs(t[1] - ref))
        best_key, best_rt, best_name = cands[0]
        if abs(best_rt - ref) > cap:
            for key, rt, name in cands:
                demoted.append({'key': key, 'rt': rt, 'name': name, 'carbon': c,
                                'ref_rt': ref, 'delta': round(rt - ref, 3),
                                'by_rt': None, 'cap': cap, 'why': 'too_far'})
            continue
        anchors.append({'carbon': c, 'ref_rt': ref, 'key': best_key,
                        'rt': best_rt, 'name': best_name,
                        'delta': round(best_rt - ref, 3),
                        'warn': bool(abs(best_rt - ref) > NALKANE_RT_WARN)})
        for key, rt, name in cands[1:]:
            demoted.append({'key': key, 'rt': rt, 'name': name, 'carbon': c,
                            'ref_rt': ref, 'delta': round(rt - ref, 3),
                            'by_rt': best_rt, 'why': 'taken'})
    return demoted, anchors, unverified


def _nalkane_reason_en(d):
    """把降級代碼轉成 Excel 驗證報告用的英文說明。"""
    why = d.get('why')
    if why == 'no_ruler':
        return 'no C%d entry on the ruler' % d['carbon']
    if why == 'too_far':
        return ('%.2f min from the C%d ruler entry (%.3f), beyond the %.2f min tolerance'
                % (abs(d['delta']), d['carbon'], d['ref_rt'], d['cap']))
    if why == 'taken':
        return ('C%d already claimed by the peak at RT %.3f (closer to ruler %.3f)'
                % (d['carbon'], d['by_rt'], d['ref_rt']))
    return why or ''


def apply_nalkane_rt_ruler_rows(rows, ruler, cap=NALKANE_RT_CAP, mode='strict'):
    """
    _plan_nalkane_rt_ruler 的 list-of-dict 版本，就地改寫 rows 裡的 'pred'。
    給離線回歸腳本（R5／R10 ground truth 對帳）使用，與網站走同一套判準。
    """
    items = [(i, r.get('rt'), r.get('name'))
             for i, r in enumerate(rows) if r.get('pred') == 'n-alkane']
    demoted, anchors, unverified = _plan_nalkane_rt_ruler(items, ruler, cap, mode)
    for d in demoted:
        rows[d['key']]['pred'] = 'iso-alkane'
    return demoted, anchors, unverified


def _find_header_row(df_raw):
    """
    在原始（header=None）DataFrame 中找出含 'Peak#' 或 'Ret.Time' 的表頭列。
    找不到回傳 None（代表這個工作表不是有效的 Peak Table，呼叫端應略過）。
    若同一工作表出現多次（例如原始 GCMSsolution 匯出檔的 MS Similarity Search 區塊
    也含類似欄位），一律取「第一個」出現的列——這對應 MC Peak Table 一定排在最前面的慣例。
    """
    for i, row in df_raw.iterrows():
        vals = [str(v).strip().replace('\xa0', '') for v in row.values]
        if 'Peak#' in vals or 'Ret.Time' in vals:
            return i
    return None


def _find_table_end(df_raw, header_row_idx):
    """
    找出 MC Peak Table 的結束列（exclusive）。

    Shimadzu GCMSsolution 的原始匯出檔在 MC Peak Table 之後，還會接著
    [MS Similarity Search Results]、[MS Spectrum] 等區塊，動輒數千列。
    若不設邊界，那些列會被當成 peak 一起送進分類流程 —— 除了統計失真，
    更會觸發數千次 PubChem 查詢，直接讓遠端 worker 逾時（回傳 HTML 錯誤頁，
    前端就會出現 "Unexpected token '<' ... is not valid JSON"）。

    判定方式：Peak# 欄位必須是數字，遇到第一個非數字／空白列即視為表格結尾。
    例外：手動標記 n-alkane 標準峰時常把 Peak# 直接覆寫成 "C8"、"C16"
    （R5／R10 兩份人工分帳檔都是這樣），這種寫法不算表格結束。
    """
    header = [str(v).strip().replace('\xa0', '') for v in df_raw.iloc[header_row_idx].values]
    key_col = None
    for candidate in ('Peak#', 'Ret.Time'):
        if candidate in header:
            key_col = header.index(candidate)
            break
    if key_col is None:
        return len(df_raw)

    end = header_row_idx + 1
    for i in range(header_row_idx + 1, len(df_raw)):
        v = df_raw.iloc[i, key_col]
        if pd.isna(v):
            break
        if re.fullmatch(r'C\s*\d{1,2}', str(v).strip(), re.IGNORECASE):
            end = i + 1
            continue
        try:
            float(v)
        except (TypeError, ValueError):
            break
        end = i + 1
    return end


def _process_single_sheet(df_raw, header_row_idx, ws, budget=None, opts=None):
    """
    對單一工作表執行：清理 → 分類 → 統計 → 寫回 Excel（上色、Summary、圖表、驗證報告）。
    ws 是這個工作表對應的 openpyxl worksheet（屬於呼叫端已載入的同一份 wb，會被就地修改）。
    budget 是整份活頁簿共用的 PubChem 查詢時間預算（見 _PubChemBudget）。
    opts 是前端表單帶進來的分析設定（溶劑扣除方式、碳數視窗邊界）。
    回傳統計 dict（不含 'sheet'/'status'，由呼叫端 process_gcms_workbook 補上）。
    """
    opts = opts or {}
    rt_c8_min  = float(opts.get('rt_c8_min')  or RT_C8_MINUS_MAX)
    rt_c16_max = float(opts.get('rt_c16_max') or RT_C816_MAX)
    nalk_mode  = (opts.get('nalkane_rt_mode') or 'strict').lower()
    nalk_ruler = opts.get('nalkane_rt_ruler') or DEFAULT_NALKANE_RT
    nalk_cap   = float(opts.get('nalkane_rt_cap') or NALKANE_RT_CAP)

    table_end = _find_table_end(df_raw, header_row_idx)
    data = df_raw.iloc[header_row_idx:table_end].copy()
    data.columns = data.iloc[0]
    # 明確 .copy()：切完片再直接指派欄位會被 pandas 判定為 chained assignment，
    # 除了在 log 洗出大量 FutureWarning，pandas 3.0 起也會真的失效。
    data = data.iloc[1:].reset_index(drop=True).copy()
    data.columns = [str(c).strip().replace('\xa0', '') for c in data.columns]

    # 必要欄位檢查：表頭若被合併儲存格、手動改名，會在這裡直接抓不到欄位
    required_cols = {'Ret.Time', 'Conc.', 'Name'}
    missing = required_cols - set(data.columns)
    if missing:
        raise ValueError(
            f'缺少必要欄位：{"、".join(sorted(missing))}'
            '（表頭列可能被合併儲存格、手動改了欄名，或欄名前後多了空白字元）'
        )

    # Step 2: 清理
    data['Ret.Time'] = pd.to_numeric(data['Ret.Time'], errors='coerce')
    data['Conc.']    = pd.to_numeric(data['Conc.'],    errors='coerce')
    data['Name']     = data['Name'].astype(str).str.strip()
    if 'Category' in data.columns:
        data['Category'] = data['Category'].astype(str).str.strip().str.replace('\xa0','',regex=False)
    else:
        data['Category'] = ''

    # 取這一列的 Peak#。人工分帳檔常把 n-alkane 標準峰的 Peak# 覆寫成 "C8"、"C16"，
    # 所以不能無條件 int()（會直接 ValueError 讓整個工作表報錯）。
    def _peak_no(i):
        if 'Peak#' in data.columns and pd.notna(data.at[i, 'Peak#']):
            try:
                return int(float(data.at[i, 'Peak#']))
            except (TypeError, ValueError):
                return str(data.at[i, 'Peak#']).strip()
        return int(i) + 1

    # Step 2.5: 標出溶劑峰（在分類之前，才不會對溶劑峰浪費 PubChem 查詢）
    solvent_idx   = _resolve_solvent_peaks(data, {**opts, 'rt_c8_min': rt_c8_min})
    is_solvent    = data.index.isin(solvent_idx)
    solvent_total = round(float(data.loc[is_solvent, 'Conc.'].sum()), 2)
    solvent_peaks = [
        {'peak':  _peak_no(i),
         'rt':    float(data.at[i, 'Ret.Time']),
         'name':  str(data.at[i, 'Name']),
         'conc':  float(data.at[i, 'Conc.'])}
        for i in solvent_idx
    ]

    # Step 3: 自動分類（溶劑峰不分類，直接掛 solvent 標記）
    mask_empty = ((data['Category'].isin(['nan','','NaN']) | data['Category'].isna())
                  & ~is_solvent)
    pending = data.loc[mask_empty, 'Name']

    # 先用關鍵字規則篩一輪，把判不出來的名稱集中起來並行查 PubChem，
    # 避免在下面的 apply 裡一個一個等網路（那正是 Render 逾時的原因）。
    unresolved = [nm for nm in pending.unique() if not _classify_keyword(str(nm))[1]]
    _prefetch_pubchem(unresolved, budget)

    data.loc[mask_empty, 'Category'] = pending.apply(
        lambda nm: _pubchem_fallback(nm, budget)
    )
    auto_classified_count = int(mask_empty.sum())
    data.loc[is_solvent, 'Category'] = SOLVENT_CATEGORY

    # Step 3.5: n-alkane 滯留時間校驗（見 _plan_nalkane_rt_ruler）。
    #
    # 只校驗「這一輪自動分類出來的 n-alkane」——使用者手動填在 Category 欄的
    # 值一律尊重，不覆蓋。降級後的資訊寫進驗證報告，看得到每一筆為什麼被改。
    nalkane_demoted, nalkane_anchors, nalkane_unverified = [], [], []
    if nalk_mode != 'off':
        ruler_items = [(i, data.at[i, 'Ret.Time'], data.at[i, 'Name'])
                       for i in data.index
                       if mask_empty.at[i] and data.at[i, 'Category'] == 'n-alkane']
        nalkane_demoted, nalkane_anchors, nalkane_unverified = _plan_nalkane_rt_ruler(
            ruler_items, nalk_ruler, nalk_cap, nalk_mode)
        for d in nalkane_demoted:
            data.at[d['key'], 'Category'] = 'iso-alkane'
            d['peak'] = _peak_no(d['key'])
            d['conc'] = (float(data.at[d['key'], 'Conc.'])
                         if pd.notna(data.at[d['key'], 'Conc.']) else 0.0)
        for a in nalkane_anchors:
            a['peak'] = _peak_no(a['key'])
        for u in nalkane_unverified:
            u['peak'] = _peak_no(u['key'])

    # Step 4: 計算統計（在操作 Excel 前完成）
    #
    # 所有統計都只算「產物峰」——溶劑峰整筆排除在分子與分母之外。
    # 排除之後剩下的峰加起來不再是 100%，因此再正規化回 100%
    # （等同手動算法裡的 "without solvent" 那一列）。
    prod = data[~is_solvent]

    mask_c8m  = prod['Ret.Time'] < rt_c8_min
    mask_c816 = (prod['Ret.Time'] >= rt_c8_min) & (prod['Ret.Time'] <= rt_c16_max)
    mask_c16p = prod['Ret.Time'] > rt_c16_max

    c8minus_raw    = round(float(prod.loc[mask_c8m,  'Conc.'].sum()), 2)
    c16plus_raw    = round(float(prod.loc[mask_c16p, 'Conc.'].sum()), 2)
    filt816        = prod[mask_c816]
    c816_total_raw = round(float(filt816['Conc.'].sum()), 2)
    sums_816_raw   = filt816.groupby('Category', dropna=True)['Conc.'].sum()

    # 正規化因子：以「扣掉溶劑後的全部產物峰」為 100%
    prod_sum = float(prod['Conc.'].sum())
    factor   = (100.0 / prod_sum) if prod_sum > 0 else 1.0

    c8minus    = round(c8minus_raw    * factor, 2)
    c816_total = round(c816_total_raw * factor, 2)
    c16plus    = round(c16plus_raw    * factor, 2)
    sums_816   = (sums_816_raw * factor).round(2)

    n_val     = float(sums_816.get('n-alkane',   0.0))
    iso_val   = float(sums_816.get('iso-alkane', 0.0))
    # n-alkane 為 0 時 i/n 是未定義，不是 0。
    # 純芳香化產物（ZSM-5 之類）真的會出現這種情形，報 0.00 會被讀成
    # 「完全沒有支鏈」，與事實相反。
    i_n_ratio = round(iso_val / n_val, 2) if n_val > 0 else None
    i_n_display = i_n_ratio if i_n_ratio is not None else 'n/a'

    # 類別 × 碳數視窗矩陣。舊版只輸出 C8-16 一段的類別明細，
    # C8- 與 C16+ 只有總量、沒有組成，重質產物（>C16）的鏈長選擇性無從判讀。
    window_of = pd.Series('C16+', index=prod.index)
    window_of[mask_c8m]  = 'C8-'
    window_of[mask_c816] = 'C8-16'
    matrix = {}
    for win in ('C8-', 'C8-16', 'C16+'):
        sub = prod[window_of == win]
        g   = (sub.groupby('Category', dropna=True)['Conc.'].sum() * factor).round(2)
        matrix[win] = {k: float(v) for k, v in g.items()}

    grand_total  = round(c8minus_raw + c816_total_raw + c16plus_raw, 2)
    all_conc_sum = round(float(data['Conc.'].sum()), 2)
    error        = abs(grand_total + solvent_total - all_conc_sum)
    passed       = error < 0.015
    others_816   = filt816[filt816['Category'] == 'others'][['Ret.Time','Name','Conc.']].to_dict('records')

    # 名稱碳數 vs 滯留時間一致性稽核（不改分類，只出警示清單）。
    # 分類完全建立在 NIST 比對出來的名稱上，名稱錯了分類一定跟著錯；
    # 碳數與滯留位置對不上是「圖庫比對品質差」最容易自動抓到的信號。
    carbon_flags = []
    for i in prod.index:
        rt = prod.at[i, 'Ret.Time']
        if pd.isna(rt) or prod.at[i, 'Category'] not in _CARBON_CHECK_CATEGORIES:
            continue
        expected = _carbon_from_rt(float(rt), rt_c8_min, rt_c16_max)
        named    = _name_carbon_count(prod.at[i, 'Name'])
        if expected is None or named is None:
            continue
        if abs(named - expected) >= CARBON_MISMATCH_TOL:
            carbon_flags.append({
                'peak':     _peak_no(i),
                'rt':       float(rt),
                'name':     str(prod.at[i, 'Name']),
                'conc':     float(prod.at[i, 'Conc.']),
                'named_c':  int(named),
                'rt_c':     round(expected, 1),
            })
    carbon_flag_conc = round(sum(f['conc'] for f in carbon_flags) * factor, 2)

    # Step 5: 確認是否已有 Summary 表格（ws 由呼叫端傳入，屬於已載入的工作簿）
    summary_header_row = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == 'C8-':
                summary_header_row = cell.row
                break
        if summary_header_row:
            break

    rows_inserted = 0
    if summary_header_row is None:
        insert_at = header_row_idx + 1  # 0-based → 1-based
        ws.insert_rows(insert_at, amount=4)
        rows_inserted = 4
        summary_header_row = insert_at

    # DATA_START_ROW 同步更新
    DATA_START_ROW = header_row_idx + 2 + rows_inserted

    # Step 6: 找/建 Category、Distribution 欄
    peak_header_excel_row = header_row_idx + 1 + rows_inserted
    cat_col_excel = None
    for cell in ws[peak_header_excel_row]:
        v = str(cell.value).strip().replace('\xa0','') if cell.value else ''
        if v == 'Category':
            cat_col_excel = cell.column

    if cat_col_excel is None:
        cat_col_excel = ws.max_column + 1
        ws.cell(row=peak_header_excel_row, column=cat_col_excel,
                value='Category').font = Font(bold=True)

    dist_col_excel = cat_col_excel + 1
    if ws.cell(row=peak_header_excel_row, column=dist_col_excel).value != 'Distribution':
        ws.cell(row=peak_header_excel_row, column=dist_col_excel,
                value='Distribution').font = Font(bold=True)

    # Step 7: 上色（只染 Peak Table 資料列）
    for idx, row_data in data.iterrows():
        excel_row = DATA_START_ROW + idx
        rt       = row_data['Ret.Time']
        category = row_data['Category']

        ws.cell(row=excel_row, column=cat_col_excel).value = category
        if pd.notna(rt):
            if category == SOLVENT_CATEGORY:
                dist = 'solvent (excluded)'
            elif rt < rt_c8_min:
                dist = 'C8-'
            elif rt <= rt_c16_max:
                dist = 'C8-16'
            else:
                dist = 'C16+'
            ws.cell(row=excel_row, column=dist_col_excel).value = dist

        # 溶劑列整列染灰，一眼看得出這幾筆沒有進任何統計
        if category == SOLVENT_CATEGORY:
            fill = PatternFill(start_color=COLOR_MAP[SOLVENT_CATEGORY],
                               end_color=COLOR_MAP[SOLVENT_CATEGORY], fill_type='solid')
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=excel_row, column=col_idx).fill = fill
        elif pd.notna(rt) and rt_c8_min <= rt <= rt_c16_max and category in COLOR_MAP:
            fill = PatternFill(start_color=COLOR_MAP[category],
                               end_color=COLOR_MAP[category], fill_type='solid')
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=excel_row, column=col_idx).fill = fill

    # Step 8: 寫 Summary 表格
    SUMMARY_VALUES_ROW = summary_header_row + 1
    summary_labels = ['C8-','C8-16','C16+','i/n',
                      'alkene','aromatic','cyclic','iso','n-alkane','others','poly',
                      'solvent']
    header_colors  = ['FFD966','FFD966','FFD966','FFD966',
                      'ADD8E6','92D050','FFD700','FF6600','FF8C00','FA8072','C5B0D5',
                      'BFBFBF']

    col_map = {}
    for cell in ws[summary_header_row]:
        if cell.value:
            col_map[str(cell.value).strip()] = cell.column

    if not col_map:
        for col_offset, (label, color) in enumerate(zip(summary_labels, header_colors)):
            cell = ws.cell(row=summary_header_row, column=col_offset+1)
            cell.value     = label
            cell.font      = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.fill      = PatternFill(start_color=color, end_color=color, fill_type='solid')
            col_map[label] = col_offset + 1

    value_map = {
        'C8-':      c8minus,
        'C8-16':    c816_total,
        'C16+':     c16plus,
        'i/n':      i_n_display,
        'alkene':   round(sums_816.get('Alkene',       0.0), 2),
        'aromatic': round(sums_816.get('Aromatic',     0.0), 2),
        'cyclic':   round(sums_816.get('cyclic',       0.0), 2),
        'iso':      round(sums_816.get('iso-alkane',   0.0), 2),
        'n-alkane': round(sums_816.get('n-alkane',     0.0), 2),
        'others':   round(sums_816.get('others',       0.0), 2),
        'poly':     round(sums_816.get('Polyaromatic', 0.0), 2),
        # solvent 是「被排除掉的量」，用原始 Conc%（未正規化），
        # 因為它不屬於正規化後的 100% 之內。
        'solvent':  solvent_total,
    }
    for label, val in value_map.items():
        if label in col_map:
            cell = ws.cell(row=SUMMARY_VALUES_ROW, column=col_map[label])
            cell.value         = val
            cell.number_format = '0.00'
            cell.alignment     = Alignment(horizontal='center')
            cell.font          = Font(bold=True)
            cell.fill          = PatternFill(start_color='FFF2CC',
                                             end_color='FFF2CC', fill_type='solid')

    # Step 9: BarChart（放在 Summary 右側 N 欄）
    chart_data_col = len(summary_labels) + 3
    chart_labels = ['Aromatic','Polyaromatic','n-alkane','iso-alkane','cyclic','Alkene','others']
    chart_values = [
        round(sums_816.get('Aromatic',     0.0), 2),
        round(sums_816.get('Polyaromatic', 0.0), 2),
        round(sums_816.get('n-alkane',     0.0), 2),
        round(sums_816.get('iso-alkane',   0.0), 2),
        round(sums_816.get('cyclic',       0.0), 2),
        round(sums_816.get('Alkene',       0.0), 2),
        round(sums_816.get('others',       0.0), 2),
    ]

    ws.cell(row=summary_header_row,   column=chart_data_col).value = 'Category'
    ws.cell(row=summary_header_row,   column=chart_data_col+1).value = 'Conc%'
    for i, (lbl, val) in enumerate(zip(chart_labels, chart_values)):
        ws.cell(row=summary_header_row+1+i, column=chart_data_col).value   = lbl
        ws.cell(row=summary_header_row+1+i, column=chart_data_col+1).value = val

    from openpyxl.chart.label import DataLabel, DataLabelList
    from openpyxl.chart.layout import Layout

    chart = BarChart()
    chart.type       = 'col'
    chart.grouping   = 'clustered'
    chart.title      = 'C8-16 Category Distribution (%, solvent-free basis)'
    chart.y_axis.title  = 'Conc. %'
    chart.x_axis.title  = 'Category'
    chart.width      = 22
    chart.height     = 14
    chart.y_axis.numFmt = '0.00'
    chart.y_axis.majorGridlines = None

    data_ref = Reference(ws,
                         min_col=chart_data_col+1,
                         min_row=summary_header_row,
                         max_row=summary_header_row+len(chart_labels))
    cats_ref = Reference(ws,
                         min_col=chart_data_col,
                         min_row=summary_header_row+1,
                         max_row=summary_header_row+len(chart_labels))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    bar_colors = ['92D050','C5B0D5','FF8C00','FF6600','FFD700','ADD8E6','FA8072']
    for idx, color in enumerate(bar_colors):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = color
        chart.series[0].dPt.append(pt)

    # 加 data labels（顯示數值在 bar 上方）
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal   = True
    chart.series[0].dLbls.showLegendKey = False
    chart.series[0].dLbls.showCatName   = False
    chart.series[0].dLbls.numFmt        = '0.00'

    ws.add_chart(chart, f'N{summary_header_row}')

    # Step 10: Validation Report（Peak Table 最後一列之後）
    report_start_row = DATA_START_ROW + len(data) + 2

    solvent_label = ', '.join('#%s' % pk['peak'] for pk in solvent_peaks) or '(none)'
    report_rows = [
        ('Basis', 'solvent-free, normalized to 100%'),
        ('Solvent excluded (raw conc%)', solvent_total),
        ('Solvent peaks',        solvent_label),
        ('C8 / C16 RT window',   '%s - %s min' % (rt_c8_min, rt_c16_max)),
        ('n-alkane RT check',    ('off' if nalk_mode == 'off' else
                                  '%s, tol %.2f min, ruler C%d-C%d'
                                  % (nalk_mode, nalk_cap,
                                     min(nalk_ruler), max(nalk_ruler)))),
        ('  matched n-alkanes',  len(nalkane_anchors)),
        ('  demoted to iso',     '%d peaks / %.2f conc%%'
                                 % (len(nalkane_demoted),
                                    round(sum(d['conc'] for d in nalkane_demoted)
                                          * factor, 2))),
        ('  kept unverified',    len(nalkane_unverified)),
        ('C8-  (conc%)',         c8minus),
        ('C8-16  (conc%)',       c816_total),
        ('C16+  (conc%)',        c16plus),
        ('  Aromatic',           round(sums_816.get('Aromatic',     0.0), 2)),
        ('  Polyaromatic',       round(sums_816.get('Polyaromatic', 0.0), 2)),
        ('  n-alkane',           round(sums_816.get('n-alkane',     0.0), 2)),
        ('  iso-alkane',         round(sums_816.get('iso-alkane',   0.0), 2)),
        ('  cyclic',             round(sums_816.get('cyclic',       0.0), 2)),
        ('  Alkene',             round(sums_816.get('Alkene',       0.0), 2)),
        ('  others',             round(sums_816.get('others',       0.0), 2)),
        ('  i/n ratio',          i_n_display),
        ('Total check (raw sum)', grand_total),
        ('All Conc. sum',        all_conc_sum),
        ('delta error',          round(error, 4)),
        ('PASS / FAIL',          'PASS' if passed else f'FAIL (delta={round(error,4)})'),
        ('Auto-classified rows', auto_classified_count),
    ]

    # 先把需要的列數插進去，不要直接往下寫。
    # Shimadzu 原始匯出檔在 Peak Table 後面還接著 [MS Chromatogram]（動輒數千列），
    # 直接覆寫會把原始層析圖資料吃掉一段 —— 舊版就是這樣，只是報告短、破壞範圍小。
    cats_order = ['n-alkane','iso-alkane','Alkene','cyclic','Aromatic','Polyaromatic','others']
    rows_needed = (1 + len(report_rows) + 1              # 標題 + 統計 + 空列
                   + 2 + len(cats_order) + 1)            # 類別 x 視窗矩陣
    if nalkane_anchors:
        rows_needed += 2 + len(nalkane_anchors) + 1
    if nalkane_demoted:
        rows_needed += 2 + len(nalkane_demoted) + 1
    if others_816:
        rows_needed += 1 + len(others_816) + 1
    if carbon_flags:
        rows_needed += 2 + len(carbon_flags) + 1
    ws.insert_rows(report_start_row, amount=rows_needed)

    title_cell = ws.cell(row=report_start_row, column=1)
    title_cell.value = 'Validation Report'
    title_cell.fill  = PatternFill(start_color='2B6CB0', end_color='2B6CB0', fill_type='solid')
    title_cell.font  = Font(bold=True, color='FFFFFF', size=11)

    for offset, (label, value) in enumerate(report_rows):
        r = report_start_row + 1 + offset
        ws.cell(row=r, column=1).value = label
        ws.cell(row=r, column=2).value = value
        ws.cell(row=r, column=1).font  = Font(size=10)
        ws.cell(row=r, column=2).font  = Font(bold=True, size=10)
        if label == 'PASS / FAIL':
            bg = 'C6EFCE' if passed else 'FFC7CE'
            ws.cell(row=r, column=2).fill = PatternFill(
                start_color=bg, end_color=bg, fill_type='solid')

    next_row = report_start_row + 1 + len(report_rows) + 1

    # 類別 x 碳數視窗矩陣。舊版只輸出 C8-16 一段的類別明細，
    # C8- 與 C16+ 只有總量、沒有組成，重質產物的鏈長選擇性無從判讀。
    ws.cell(row=next_row, column=1).value = 'Category x carbon window (conc%, solvent-free)'
    ws.cell(row=next_row, column=1).font  = Font(bold=True)
    for j, win in enumerate(('C8-', 'C8-16', 'C16+')):
        c = ws.cell(row=next_row + 1, column=2 + j)
        c.value = win
        c.font  = Font(bold=True)
    for i, cat in enumerate(cats_order):
        ws.cell(row=next_row + 2 + i, column=1).value = cat
        for j, win in enumerate(('C8-', 'C8-16', 'C16+')):
            ws.cell(row=next_row + 2 + i, column=2 + j).value = matrix[win].get(cat, 0.0)
    next_row += 2 + len(cats_order) + 1

    # n-alkane RT 標定結果：每個碳數對到哪一根峰、離標尺多遠。
    # delta 欄是整份分析裡最有用的健檢——同一次分析的偏移應該平滑變化，
    # 某個碳數突然跳掉就代表那根峰的圖庫比對（或標尺本身）有問題。
    if nalkane_anchors:
        ws.cell(row=next_row, column=1).value = 'n-alkane RT anchors (ruler vs measured)'
        ws.cell(row=next_row, column=1).font  = Font(bold=True, color='2B6CB0')
        for j, lbl in enumerate(('Carbon', 'Ruler RT', 'Peak#', 'Measured RT', 'delta', 'Name')):
            ws.cell(row=next_row + 1, column=1 + j).value = lbl
            ws.cell(row=next_row + 1, column=1 + j).font  = Font(bold=True, size=9)
        for i, a in enumerate(nalkane_anchors):
            r = next_row + 2 + i
            for j, v in enumerate(('C%d' % a['carbon'], a['ref_rt'], a['peak'],
                                   a['rt'], a['delta'], a['name'])):
                ws.cell(row=r, column=1 + j).value = v
            if a['warn']:
                ws.cell(row=r, column=5).fill = PatternFill(
                    start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        next_row += 2 + len(nalkane_anchors) + 1

    if nalkane_demoted:
        ws.cell(row=next_row, column=1).value = (
            'Demoted n-alkane -> iso-alkane by RT ruler (%d peaks)' % len(nalkane_demoted))
        ws.cell(row=next_row, column=1).font = Font(bold=True, color='C05621')
        for j, lbl in enumerate(('Peak#', 'Ret.Time', 'Conc.', 'Name', 'Reason')):
            ws.cell(row=next_row + 1, column=1 + j).value = lbl
            ws.cell(row=next_row + 1, column=1 + j).font  = Font(bold=True, size=9)
        for i, d in enumerate(nalkane_demoted):
            r = next_row + 2 + i
            for j, v in enumerate((d['peak'], d['rt'], d['conc'], d['name'],
                                   _nalkane_reason_en(d))):
                ws.cell(row=r, column=1 + j).value = v
        next_row += 2 + len(nalkane_demoted) + 1

    if others_816:
        ws.cell(row=next_row, column=1).value = 'Unclassified (others) in C8-16'
        ws.cell(row=next_row, column=1).font  = Font(bold=True, color='FF0000')
        for i, rec in enumerate(others_816):
            r = next_row + 1 + i
            ws.cell(row=r, column=1).value = rec.get('Name', '')
            ws.cell(row=r, column=2).value = rec.get('Ret.Time', '')
            ws.cell(row=r, column=3).value = rec.get('Conc.', '')
        next_row += 1 + len(others_816) + 1

    if carbon_flags:
        ws.cell(row=next_row, column=1).value = (
            'Library-ID warning: name carbon number vs retention time '
            '(|delta| >= %d C)' % CARBON_MISMATCH_TOL)
        ws.cell(row=next_row, column=1).font = Font(bold=True, color='C05621')
        for j, lbl in enumerate(('Peak#', 'Ret.Time', 'Conc.', 'C from name', 'C from RT', 'Name')):
            ws.cell(row=next_row + 1, column=1 + j).value = lbl
            ws.cell(row=next_row + 1, column=1 + j).font  = Font(bold=True, size=9)
        for i, f in enumerate(carbon_flags):
            r = next_row + 2 + i
            for j, v in enumerate((f['peak'], f['rt'], f['conc'],
                                   f['named_c'], f['rt_c'], f['name'])):
                ws.cell(row=r, column=1 + j).value = v

    # 明確轉成原生 Python 型別（pandas/numpy 運算結果是 numpy.float64／numpy.bool_，
    # 標準 json 模組無法序列化，之後要塞進 response header 的 JSON payload會直接炸掉）
    validation_report = {
        'c8minus':         float(c8minus),
        'c816_total':      float(c816_total),
        'c16plus':         float(c16plus),
        'grand_total':     float(grand_total),
        'all_conc_sum':    float(all_conc_sum),
        'error':           float(round(error, 4)),
        'passed':          bool(passed),
        'auto_classified': auto_classified_count,
        'sums_816':        {k: float(round(v, 2)) for k, v in sums_816.items()},
        'i_n_ratio':       (float(i_n_ratio) if i_n_ratio is not None else None),
        'solvent_total':   float(solvent_total),
        'solvent_peaks':   solvent_peaks,
        'rt_window':       [rt_c8_min, rt_c16_max],
        'matrix':          matrix,
        'carbon_flags':      carbon_flags[:20],
        'carbon_flag_count': len(carbon_flags),
        'carbon_flag_conc':  float(carbon_flag_conc),
        'nalkane_rt_mode':   nalk_mode,
        'nalkane_rt_warn':   NALKANE_RT_WARN,
        # 欄位名刻意縮到最短：整包 JSON 要塞進 HTTP header，而 header 有
        # 8KB 上限（gunicorn／反向代理），超過就得退回精簡版、連類別佔比都看不到。
        'nalkane_anchors':   [{'c': a['carbon'], 'ref': a['ref_rt'],
                               'pk': a['peak'], 'rt': a['rt'],
                               'd': a['delta'], 'w': a['warn']}
                              for a in nalkane_anchors],
        'nalkane_demoted':   [{'pk': d['peak'], 'rt': d['rt'],
                               'conc': round(d['conc'] * factor, 2),
                               'nm': d['name'], 'why': d['why'],
                               'c': d.get('carbon'), 'ref': d.get('ref_rt'),
                               'd': d.get('delta'), 'by': d.get('by_rt'),
                               'cap': d.get('cap')}
                              for d in nalkane_demoted[:12]],
        'nalkane_demoted_count': len(nalkane_demoted),
        'nalkane_demoted_conc':  float(round(
            sum(d['conc'] for d in nalkane_demoted) * factor, 2)),
    }
    return validation_report


def process_gcms_workbook(file_bytes: bytes, original_filename: str, opts=None):
    """
    自動掃描整份活頁簿的「每一個」工作表：
    只要該工作表含有效 Peak Table 表頭（'Peak#' 或 'Ret.Time'），就自動偵測並分類；
    不含的工作表視為無關資料（例如備註頁、其他分析結果），直接略過、不影響其他工作表。
    回傳 (result_bytes, sheet_reports)：
      - result_bytes：同一份 xlsx，每個成功分類的工作表已就地寫入 Category/Distribution/
        Summary/圖表/驗證報告；被略過或出錯的工作表原封不動。
      - sheet_reports：list[dict]，每個工作表一筆，含 'sheet'、'status'
        ('ok' | 'skipped' | 'error') 與對應的統計數字或原因說明。
    """
    # 一次解析整份活頁簿（sheet_name=None 回傳 {sheet: DataFrame}）。
    # 先前是在迴圈裡逐一 read_excel，等於每個工作表都把整個檔案重新 parse 一遍。
    try:
        all_sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)
        wb = load_workbook(io.BytesIO(file_bytes))
    except Exception:
        raise ValueError(
            '這個檔案無法被當成 Excel 開啟。可能是檔案損毀，或副檔名雖然是 .xlsx '
            '但內容其實是舊版 .xls／.csv／其他格式。請在 Excel 中開啟後，'
            '用「另存新檔 → Excel 活頁簿 (.xlsx)」重新輸出再上傳。'
        )

    budget = _PubChemBudget()
    sheet_reports = []
    any_ok = False

    for sheet_name, df_raw in all_sheets.items():
        header_row_idx = _find_header_row(df_raw)
        if header_row_idx is None:
            sheet_reports.append({
                'sheet': sheet_name, 'status': 'skipped',
                'reason': '找不到 Peak Table 表頭（需含「Peak#」或「Ret.Time」欄位），已略過此工作表',
            })
            continue

        try:
            report = _process_single_sheet(df_raw, header_row_idx, wb[sheet_name], budget, opts)
            report['sheet']  = sheet_name
            report['status'] = 'ok'
            sheet_reports.append(report)
            any_ok = True
        except Exception as e:
            sheet_reports.append({'sheet': sheet_name, 'status': 'error', 'reason': str(e)})

    if not any_ok:
        raise ValueError(
            '整份活頁簿中沒有任何工作表含有效的 Peak Table'
            '（需要「Peak#」或「Ret.Time」表頭），請確認上傳的是 Shimadzu GCMSsolution 匯出格式'
        )

    if budget.skipped:
        # 讓使用者知道有部分化合物因時間預算未做 PubChem 二次判定（仍留在 others）
        for r in sheet_reports:
            if r.get('status') == 'ok':
                r['pubchem_skipped'] = budget.skipped
                break

    # wb.save() 要把整份活頁簿序列化成 XML，是整個流程記憶體的最高點；
    # 先釋放已用不到的 DataFrame，壓低峰值（Render 免費方案只有 512MB）。
    all_sheets.clear()
    del all_sheets

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read(), sheet_reports


# ── 錯誤一律回 JSON ────────────────────────────────────────────────────
# 預設 Flask/Werkzeug 的錯誤頁是 HTML，前端 res.json() 解析時會炸出
# "Unexpected token '<' ... is not valid JSON"，讓真正的錯誤原因完全看不到。
@app.errorhandler(HTTPException)
def _handle_http_exception(e):
    if e.code == 413:
        msg = '檔案超過 16MB 上限，請先刪掉用不到的工作表或改上傳精簡後的 Peak Table'
    else:
        msg = f'{e.code} {e.name}：{e.description}'
    return jsonify({'error': msg}), e.code


@app.errorhandler(Exception)
def _handle_unexpected(e):
    app.logger.exception('unhandled error in request')
    return jsonify({'error': f'伺服器未預期錯誤（{type(e).__name__}）：{e}'}), 500


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/classify', methods=['POST'])
def classify():
    if 'file' not in request.files:
        return jsonify({'error': '沒有收到檔案，請重新選擇後再送出'}), 400
    f = request.files['file']
    if not f.filename.lower().endswith('.xlsx'):
        return jsonify({
            'error': f'只接受 .xlsx 格式（收到的是「{f.filename}」）。'
                     '舊版 .xls 與 .csv 不支援，請先在 Excel 另存成 .xlsx 再上傳。'
        }), 400

    file_bytes = f.read()
    original_filename = secure_filename(f.filename)

    # 分析設定：溶劑扣除方式與碳數視窗邊界。
    # 空字串一律當成「沒填」→ 用預設值，不要讓 float('') 直接炸掉。
    def _num(field, default):
        raw = (request.form.get(field) or '').strip()
        if not raw:
            return default
        try:
            return float(raw)
        except ValueError:
            return default

    opts = {
        'solvent_mode':     (request.form.get('solvent_mode') or 'auto').strip().lower(),
        'solvent_peaks':    request.form.get('solvent_peaks') or '',
        'solvent_min_conc': _num('solvent_min_conc', SOLVENT_MIN_CONC),
        'rt_c8_min':        _num('rt_c8_min',  RT_C8_MINUS_MAX),
        'rt_c16_max':       _num('rt_c16_max', RT_C816_MAX),
        # n-alkane 滯留時間標尺。前端傳的是自由格式文字，解析不出來就用預設表。
        'nalkane_rt_mode':  (request.form.get('nalkane_rt_mode') or 'strict').strip().lower(),
        'nalkane_rt_ruler': (_parse_rt_ruler(request.form.get('nalkane_rt_ruler'))
                             or DEFAULT_NALKANE_RT),
        'nalkane_rt_cap':   _num('nalkane_rt_cap', NALKANE_RT_CAP),
    }
    if opts['nalkane_rt_mode'] not in ('strict', 'range', 'off'):
        opts['nalkane_rt_mode'] = 'strict'

    try:
        result_bytes, sheet_reports = process_gcms_workbook(file_bytes, original_filename, opts)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    stem = original_filename.rsplit('.', 1)[0]
    out_filename = f"{stem}_classified.xlsx"

    response = send_file(
        io.BytesIO(result_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=out_filename,
    )
    # 多工作表結果以 JSON（base64 編碼避免 header 特殊字元問題）夾帶在自訂 header 裡，
    # 前端解碼後逐一渲染每個工作表的統計卡片。
    payload = json.dumps(sheet_reports, ensure_ascii=False)
    encoded = base64.b64encode(payload.encode('utf-8')).decode('ascii')

    # HTTP header 有大小限制（gunicorn / 反向代理常見上限 8190 bytes）。
    # 工作表很多時整包 JSON 會超過而讓整個回應被擋掉，這時要逐級瘦身。
    #
    # 舊版只有「完整」與「只剩 sheet 名稱與狀態」兩段，落差太大：單一工作表
    # 只要明細清單稍長就整個掉到最精簡版，連七類佔比與 i/n 都看不到。
    # 改成逐級丟：先丟最貴的逐筆明細（Excel 檔裡本來就有完整清單），
    # 再丟摘要層，最後才退到只剩狀態。
    DETAIL_KEYS = ('nalkane_demoted', 'carbon_flags', 'nalkane_anchors',
                   'solvent_peaks', 'matrix')
    SUMMARY_KEYS = ('sheet', 'status', 'c8minus', 'c816_total', 'c16plus',
                    'passed', 'error', 'auto_classified', 'i_n_ratio',
                    'solvent_total', 'carbon_flag_count', 'carbon_flag_conc',
                    'nalkane_rt_mode', 'nalkane_rt_warn',
                    'nalkane_demoted_count', 'nalkane_demoted_conc', 'sums_816')

    def _encode(reports):
        return base64.b64encode(
            json.dumps(reports, ensure_ascii=False).encode('utf-8')).decode('ascii')

    def _drop(reports, keys):
        out = []
        for r in reports:
            if r.get('status') != 'ok':
                out.append({'sheet': r.get('sheet'), 'status': r.get('status'),
                            'reason': str(r.get('reason', ''))[:120]})
            else:
                out.append({k: v for k, v in r.items() if k not in keys})
        return out

    def _keep(reports, keys):
        out = []
        for r in reports:
            if r.get('status') != 'ok':
                out.append({'sheet': r.get('sheet'), 'status': r.get('status'),
                            'reason': str(r.get('reason', ''))[:120]})
            else:
                out.append({k: r[k] for k in keys if k in r})
        return out

    for stage in (lambda: sheet_reports,
                  lambda: _drop(sheet_reports, DETAIL_KEYS[:2]),
                  lambda: _drop(sheet_reports, DETAIL_KEYS),
                  lambda: _keep(sheet_reports, SUMMARY_KEYS),
                  lambda: _keep(sheet_reports, ('sheet', 'status'))):
        encoded = _encode(stage())
        if len(encoded) <= 6000:
            break

    response.headers['X-Sheet-Reports'] = encoded
    response.headers['Access-Control-Expose-Headers'] = 'X-Sheet-Reports'
    return response


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

