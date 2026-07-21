"""
GC-MS FID 成分分類網站
自動對 Shimadzu GC-MS 輸出的 xlsx 執行成分分類、顏色標記、統計加總
"""

import io
import os
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
}

RT_C8_MINUS_MAX = 2.178
RT_C816_MAX     = 17.6



# ── PubChem Hybrid 查詢（僅對 others 觸發）────────────────────────────
_pubchem_cache: dict = {}   # in-memory cache，每次重啟清空
PUBCHEM_CACHE_MAX = 5000    # 上限，避免長時間執行後無限成長吃光記憶體


def _cache_put(key: str, value: str) -> None:
    """寫入查詢快取；超過上限就整批清掉重來（用量小，不值得為此引入 LRU）。"""
    if len(_pubchem_cache) >= PUBCHEM_CACHE_MAX:
        _pubchem_cache.clear()
    _pubchem_cache[key] = value

PUBCHEM_HTTP_TIMEOUT   = 3     # 單次查詢逾時（秒）
PUBCHEM_MAX_WORKERS    = 8     # 並行查詢數（純網路等待，用執行緒即可）
# 每次上傳的 PubChem 查詢總時間上限（秒）。
# 必須明顯小於 gunicorn 的 --timeout（Render 預設只有 30 秒），
# 讓後面的寫檔（wb.save 是最慢的一步）與回傳都還有餘裕。
# 預算用完後剩下的化合物一律留在 others，並在頁面上回報略過筆數。
PUBCHEM_TIME_BUDGET    = 12


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

    try:
        encoded = urllib.parse.quote(key)
        url = (f'https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/'
               f'{encoded}/property/MolecularFormula,IUPACName/JSON')
        with urllib.request.urlopen(url, timeout=PUBCHEM_HTTP_TIMEOUT) as r:
            data = _json.loads(r.read())
        props = data['PropertyTable']['Properties'][0]
        formula = props.get('MolecularFormula', '')
        iupac   = props.get('IUPACName', '').upper()
        result  = _classify_by_formula(formula, iupac, key)
    except Exception:
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
    m = re.match(r'^C(\d+)H(\d+)((?:O\d*)?(?:N\d*)?)$', formula)
    if not m:
        return 'others'

    c = int(m.group(1))
    h = int(m.group(2))
    hetero = m.group(3)   # 含 O/N → 雜原子化合物（脂肪醇、醛、酮、酯、胺）

    # 部分氫化的萘/茚 = 單芳香環 → Aromatic（與 auto_classify 的環數慣例一致）
    if (('NAPHTHALENE' in iupac and ('TETRAHYDRO' in iupac or 'DIHYDRO' in iupac))
            or any(x in iupac for x in ['TETRALIN','INDANE','INDENE','INDACENE'])):
        return 'Aromatic'

    # 多環芳烴（≥2 個芳香環）
    if any(x in iupac for x in ['NAPHTHALENE','PYRENE','FLUORENE',
                                  'ANTHRACENE','PHENANTHRENE','BIPHENYL',
                                  'AZULENE','FLUORANTHENE','PERYLENE','CHRYSENE']):
        return 'Polyaromatic'

    # 單環芳烴（苯環 CₙH₂ₙ₋₆）
    if h == 2 * c - 6 or any(x in iupac for x in ['BENZENE','TOLUENE',
                                                     'XYLENE','PHENYL',
                                                     'STYRENE','CUMENE']):
        return 'Aromatic'

    # 帶 O/N 的非芳香化合物 → 氧化物 / 含雜原子，非純碳氫類 → others
    # （修正：脂肪醇 CₙH₂ₙ₊₂O 會被誤判成 n-alkane，因為過去只看 C/H 忽略了 O）
    if hetero:
        return 'others'

    # Cycloalkane CₙH₂ₙ + cycl in name
    if h == 2 * c and 'CYCL' in iupac:
        return 'cyclic'

    # Alkene CₙH₂ₙ（不含 cyclo）
    if h == 2 * c:
        return 'Alkene'

    # Alkane CₙH₂ₙ₊₂
    if h == 2 * c + 2:
        if 'METHYL' in iupac or 'ETHYL' in iupac:
            # 直鏈 ethyl ≠ iso，再確認是否 branched（iupac 已 .upper()，故正則需忽略大小寫）
            if re.search(r'\d+-methyl|\d+-ethyl|\d+-propyl', iupac, re.IGNORECASE):
                return 'iso-alkane'
        return 'n-alkane'

    return 'others'

def auto_classify(name: str) -> str:
    n = name.upper().strip()

    # ── 1. 部分氫化的萘/茚 = 只剩「一個」芳香環 → 單環 Aromatic ──────────
    #    tetralin(四氫萘)、dihydronaphthalene、indane、indene、indacene。
    #    採「芳香環數目」慣例:此類只有 1 個芳香環,不列 Polyaromatic(PAH 需 ≥2 芳香環)。
    #    相容 CAS 倒裝 "NAPHTHALENE, 1,2,3,4-TETRAHYDRO-" 與正寫 "...TETRAHYDRONAPHTHALENE"。
    if 'NAPHTHALENE' in n and ('DECAHYDRO' in n or 'DECALIN' in n):
        return 'cyclic'                       # 全飽和 decalin → 環烷
    if (('NAPHTHALENE' in n and ('TETRAHYDRO' in n or 'DIHYDRO' in n))
            or 'TETRALIN' in n
            or any(x in n for x in ['INDENE', 'INDAN', 'INDACENE'])):
        return 'Aromatic'

    # ── 2. 真正多環芳烴(≥2 個芳香環)→ Polyaromatic ────────────────────
    if any(x in n for x in ['NAPHTHALENE','ANTHRACENE','PHENANTHRENE',
                              'PYRENE','FLUORENE','BIPHENYL','ACENAPHTHY',
                              'AZULENE','FLUORANTHENE','PERYLENE','CHRYSENE']):
        return 'Polyaromatic'

    # ── 3. 單環芳烴(苯環)→ Aromatic ──────────────────────────────────
    if any(x in n for x in ['BENZENE','TOLUENE','XYLENE','STYRENE',
                              'PHENYL','CUMENE','CYMENE']):
        return 'Aromatic'
    if any(x in n for x in ['CYCLOPENTANE','CYCLOHEXANE','CYCLOPROPANE',
                              'CYCLOBUTANE','CYCLOHEPTANE','CYCLOOCTANE']):
        return 'cyclic'
    if '-ENE' in n or any(x in n for x in ['PROPENE','BUTENE','PENTENE',
                                             'HEXENE','HEPTENE','OCTENE',
                                             'NONENE','DECENE']):
        return 'Alkene'
    straight = ['HEXANE','HEPTANE','OCTANE','NONANE','DECANE',
                'UNDECANE','DODECANE','TRIDECANE','TETRADECANE',
                'PENTADECANE','HEXADECANE','HEPTADECANE','OCTADECANE']
    if any(x in n for x in straight):
        if '2-METHYL' in n or '3-METHYL' in n:
            return 'iso-alkane'
        return 'n-alkane'
    if any(x in n for x in ['ISOPENTANE','ISOBUTANE','2-METHYL','3-METHYL']):
        if 'ANE' in n:
            return 'iso-alkane'
    return 'others'


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
    先 auto_classify，若結果是 others 才查 PubChem。
    已快取的名稱直接取用（不佔預算）；未快取且預算已用罄時，保持 'others' 不再發出網路查詢。
    """
    result = auto_classify(name)
    if result != 'others':
        return result

    key = name.upper().strip()
    if key in _pubchem_cache:
        return _pubchem_cache[key]

    if budget is not None and not budget.allows_lookup():
        return 'others'

    return _pubchem_lookup(name)


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
        try:
            float(v)
        except (TypeError, ValueError):
            break
        end = i + 1
    return end


def _process_single_sheet(df_raw, header_row_idx, ws, budget=None):
    """
    對單一工作表執行：清理 → 分類 → 統計 → 寫回 Excel（上色、Summary、圖表、驗證報告）。
    ws 是這個工作表對應的 openpyxl worksheet（屬於呼叫端已載入的同一份 wb，會被就地修改）。
    budget 是整份活頁簿共用的 PubChem 查詢時間預算（見 _PubChemBudget）。
    回傳統計 dict（不含 'sheet'/'status'，由呼叫端 process_gcms_workbook 補上）。
    """
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

    # Step 3: 自動分類
    mask_empty = data['Category'].isin(['nan','','NaN']) | data['Category'].isna()
    pending = data.loc[mask_empty, 'Name']

    # 先用關鍵字規則篩一輪，把判不出來的名稱集中起來並行查 PubChem，
    # 避免在下面的 apply 裡一個一個等網路（那正是 Render 逾時的原因）。
    unresolved = [nm for nm in pending.unique() if auto_classify(str(nm)) == 'others']
    _prefetch_pubchem(unresolved, budget)

    data.loc[mask_empty, 'Category'] = pending.apply(
        lambda nm: _pubchem_fallback(nm, budget)
    )
    auto_classified_count = int(mask_empty.sum())

    # Step 4: 計算統計（在操作 Excel 前完成）
    mask_c8m  = data['Ret.Time'] < RT_C8_MINUS_MAX
    mask_c816 = (data['Ret.Time'] >= RT_C8_MINUS_MAX) & (data['Ret.Time'] <= RT_C816_MAX)
    mask_c16p = data['Ret.Time'] > RT_C816_MAX

    c8minus    = round(data.loc[mask_c8m,  'Conc.'].sum(), 2)
    c16plus    = round(data.loc[mask_c16p, 'Conc.'].sum(), 2)
    filt816    = data[mask_c816]
    c816_total = round(filt816['Conc.'].sum(), 2)
    sums_816   = filt816.groupby('Category', dropna=True)['Conc.'].sum()

    n_val     = sums_816.get('n-alkane',   0.0)
    iso_val   = sums_816.get('iso-alkane', 0.0)
    i_n_ratio = round(iso_val / n_val, 2) if n_val > 0 else 0.0

    grand_total  = c8minus + c816_total + c16plus
    all_conc_sum = round(data['Conc.'].sum(), 2)
    error        = abs(grand_total - all_conc_sum)
    passed       = error < 0.01
    others_816   = filt816[filt816['Category']=='others'][['Ret.Time','Name','Conc.']].to_dict('records')

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
            if rt < RT_C8_MINUS_MAX:
                dist = 'C8-'
            elif rt <= RT_C816_MAX:
                dist = 'C8-16'
            else:
                dist = 'C16+'
            ws.cell(row=excel_row, column=dist_col_excel).value = dist

        if pd.notna(rt) and RT_C8_MINUS_MAX <= rt <= RT_C816_MAX and category in COLOR_MAP:
            fill = PatternFill(start_color=COLOR_MAP[category],
                               end_color=COLOR_MAP[category], fill_type='solid')
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=excel_row, column=col_idx).fill = fill

    # Step 8: 寫 Summary 表格
    SUMMARY_VALUES_ROW = summary_header_row + 1
    summary_labels = ['C8-','C8-16','C16+','i/n',
                      'alkene','aromatic','cyclic','iso','n-alkane','others','poly']
    header_colors  = ['FFD966','FFD966','FFD966','FFD966',
                      'ADD8E6','92D050','FFD700','FF6600','FF8C00','FA8072','C5B0D5']

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
        'i/n':      i_n_ratio,
        'alkene':   round(sums_816.get('Alkene',       0.0), 2),
        'aromatic': round(sums_816.get('Aromatic',     0.0), 2),
        'cyclic':   round(sums_816.get('cyclic',       0.0), 2),
        'iso':      round(sums_816.get('iso-alkane',   0.0), 2),
        'n-alkane': round(sums_816.get('n-alkane',     0.0), 2),
        'others':   round(sums_816.get('others',       0.0), 2),
        'poly':     round(sums_816.get('Polyaromatic', 0.0), 2),
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
    chart.title      = 'C8-16 Category Distribution (%)'
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

    title_cell = ws.cell(row=report_start_row, column=1)
    title_cell.value = 'Validation Report'
    title_cell.fill  = PatternFill(start_color='2B6CB0', end_color='2B6CB0', fill_type='solid')
    title_cell.font  = Font(bold=True, color='FFFFFF', size=11)

    report_rows = [
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
        ('Total check (sum)',    grand_total),
        ('All Conc. sum',        all_conc_sum),
        ('delta error',          round(error, 4)),
        ('PASS / FAIL',          'PASS' if passed else f'FAIL (delta={round(error,4)})'),
        ('Auto-classified rows', auto_classified_count),
    ]

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

    if others_816:
        others_start = report_start_row + 1 + len(report_rows) + 1
        ws.cell(row=others_start, column=1).value = 'Unclassified (others) in C8-16'
        ws.cell(row=others_start, column=1).font  = Font(bold=True, color='FF0000')
        for i, rec in enumerate(others_816):
            r = others_start + 1 + i
            ws.cell(row=r, column=1).value = rec.get('Name', '')
            ws.cell(row=r, column=2).value = rec.get('Ret.Time', '')
            ws.cell(row=r, column=3).value = rec.get('Conc.', '')

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
    }
    return validation_report


def process_gcms_workbook(file_bytes: bytes, original_filename: str):
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
            report = _process_single_sheet(df_raw, header_row_idx, wb[sheet_name], budget)
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

    try:
        result_bytes, sheet_reports = process_gcms_workbook(file_bytes, original_filename)
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
    # 工作表很多時整包 JSON 可能超過而讓整個回應被擋掉，這時改送精簡版：
    # 只保留 sheet 名稱與狀態，捨棄各類別明細（Excel 檔內仍有完整結果）。
    if len(encoded) > 6000:
        slim = []
        for r in sheet_reports:
            if r.get('status') == 'ok':
                slim.append({k: r[k] for k in
                             ('sheet', 'status', 'c8minus', 'c816_total', 'c16plus',
                              'passed', 'error', 'auto_classified') if k in r})
            else:
                slim.append({'sheet': r.get('sheet'), 'status': r.get('status'),
                             'reason': str(r.get('reason', ''))[:120]})
        slim_payload = json.dumps(slim, ensure_ascii=False)
        encoded = base64.b64encode(slim_payload.encode('utf-8')).decode('ascii')

    response.headers['X-Sheet-Reports'] = encoded
    response.headers['Access-Control-Expose-Headers'] = 'X-Sheet-Reports'
    return response


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

