"""
GC-MS FID 成分分類網站
自動對 Shimadzu GC-MS 輸出的 xlsx 執行成分分類、顏色標記、統計加總
"""

import io
import os
import pandas as pd
from flask import Flask, request, send_file, render_template, jsonify
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB 上限

# ── 顏色設定 ──────────────────────────────────────────────────────────
COLOR_MAP = {
    'n-alkane':     'FF8C00',
    'iso-alkane':   'FF6600',
    'cyclic':       'FFD700',
    'Alkene':       'ADD8E6',
    'Aromatic':     '92D050',
    'Polyaromatic': 'C5B0D5',
    'others':       'FA8072',
}

# ── RT 截斷（固定值）──────────────────────────────────────────────────
RT_C8_MINUS_MAX  = 2.178   # RT < 2.178 → C8-
RT_C816_MAX      = 17.6    # RT 2.178–17.6 → C8-16, RT > 17.6 → C16+


def auto_classify(name: str) -> str:
    """依化合物名稱自動歸類（當 Category 欄位空白時使用）"""
    n = name.upper().strip()

    if any(x in n for x in ['NAPHTHALENE', 'ANTHRACENE', 'PHENANTHRENE',
                              'PYRENE', 'FLUORENE', 'BIPHENYL', 'ACENAPHTHY',
                              'INDENE', 'TETRAHYDRONAPHTHALENE', 'TETRALIN']):
        return 'Polyaromatic'

    if any(x in n for x in ['BENZENE', 'TOLUENE', 'XYLENE', 'STYRENE',
                              'INDAN', 'PHENYL', 'CUMENE']):
        return 'Aromatic'

    if any(x in n for x in ['CYCLOPENTANE', 'CYCLOHEXANE', 'CYCLOPROPANE',
                              'CYCLOBUTANE', 'CYCLOHEPTANE', 'CYCLOOCTANE']):
        return 'cyclic'

    if '-ENE' in n or any(x in n for x in ['PROPENE', 'BUTENE', 'PENTENE',
                                             'HEXENE', 'HEPTENE', 'OCTENE',
                                             'NONENE', 'DECENE']):
        return 'Alkene'

    straight = ['HEXANE', 'HEPTANE', 'OCTANE', 'NONANE', 'DECANE',
                'UNDECANE', 'DODECANE', 'TRIDECANE', 'TETRADECANE',
                'PENTADECANE', 'HEXADECANE', 'HEPTADECANE', 'OCTADECANE']
    if any(x in n for x in straight):
        if '2-METHYL' in n or '3-METHYL' in n:
            return 'iso-alkane'
        return 'n-alkane'

    if any(x in n for x in ['ISOPENTANE', 'ISOBUTANE', '2-METHYL', '3-METHYL']):
        if 'ANE' in n:
            return 'iso-alkane'

    return 'others'


def process_gcms_file(file_bytes: bytes, original_filename: str) -> tuple[bytes, dict]:
    """
    主處理函式：讀取 xlsx → 分類 → 加顏色 → 加總 → 驗證
    回傳 (處理後的 xlsx bytes, 驗證報告 dict)
    """

    # Step 1: 動態定位 header row
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Sheet1', header=None)
    header_row_idx = None
    for i, row in df_raw.iterrows():
        vals = [str(v) for v in row.values]
        if 'Peak#' in vals or 'Ret.Time' in vals:
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("找不到 Peak Table header（含 Peak# 或 Ret.Time 的列）")

    data = df_raw.iloc[header_row_idx:].copy()
    data.columns = data.iloc[0]
    data = data.iloc[1:].reset_index(drop=True)

    # 清理欄位名稱的隱藏字元和空白（Shimadzu 常見問題）
    data.columns = [str(c).strip().replace('\xa0', '') for c in data.columns]

    # Step 2: 清理欄位內容
    data['Ret.Time'] = pd.to_numeric(data['Ret.Time'], errors='coerce')
    data['Conc.']    = pd.to_numeric(data['Conc.'],    errors='coerce')
    data['Name']     = data['Name'].astype(str).str.strip() if 'Name' in data.columns else ''

    # 若 Category 欄存在就清理，不存在就全部設空（讓自動分類接手）
    if 'Category' in data.columns:
        data['Category'] = data['Category'].astype(str).str.strip().str.replace('\xa0', '', regex=False)
    else:
        data['Category'] = ''

    # Step 3: 所有空白 / nan Category 全部自動歸類
    mask_empty = data['Category'].isin(['nan', '', 'NaN']) | data['Category'].isna()
    data.loc[mask_empty, 'Category'] = data.loc[mask_empty, 'Name'].apply(auto_classify)
    auto_classified_count = int(mask_empty.sum())

    # Step 4: 加顏色到原始 Excel
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb['Sheet1']
    DATA_START_ROW = header_row_idx + 2  # Excel 1-indexed + skip header row

    # 同時寫入 Category 欄到 Excel（若原本沒有，新增在最後一欄之後）
    # 找 Category 欄在 Excel 的位置
    cat_col_excel = None
    name_col_excel = None
    for cell in ws[header_row_idx + 1]:  # header row in Excel (1-indexed)
        v = str(cell.value).strip().replace('\xa0', '') if cell.value else ''
        if v == 'Category':
            cat_col_excel = cell.column
        if v == 'Name':
            name_col_excel = cell.column

    # 若沒有 Category 欄，在最後一欄後新增
    if cat_col_excel is None:
        cat_col_excel = ws.max_column + 1
        ws.cell(row=header_row_idx + 1, column=cat_col_excel, value='Category').font = Font(bold=True)

    # 也新增 Distribution 欄
    dist_col_excel = cat_col_excel + 1
    if ws.cell(row=header_row_idx + 1, column=dist_col_excel).value != 'Distribution':
        ws.cell(row=header_row_idx + 1, column=dist_col_excel, value='Distribution').font = Font(bold=True)

    for idx, row_data in data.iterrows():
        excel_row = DATA_START_ROW + idx
        rt        = row_data['Ret.Time']
        category  = row_data['Category']

        # 寫入 Category 和 Distribution
        ws.cell(row=excel_row, column=cat_col_excel).value = category
        if pd.notna(rt):
            if rt < RT_C8_MINUS_MAX:
                dist = 'C8-'
            elif rt <= RT_C816_MAX:
                dist = 'C8-16'
            else:
                dist = 'C16+'
            ws.cell(row=excel_row, column=dist_col_excel).value = dist

        # 對 C8-16 範圍加顏色
        if pd.notna(rt) and RT_C8_MINUS_MAX <= rt <= RT_C816_MAX and category in COLOR_MAP:
            fill = PatternFill(start_color=COLOR_MAP[category],
                               end_color=COLOR_MAP[category], fill_type='solid')
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=excel_row, column=col_idx).fill = fill

    # Step 5: 計算加總
    mask_c8m  = data['Ret.Time'] < RT_C8_MINUS_MAX
    mask_c816 = (data['Ret.Time'] >= RT_C8_MINUS_MAX) & (data['Ret.Time'] <= RT_C816_MAX)
    mask_c16p = data['Ret.Time'] > RT_C816_MAX

    c8minus    = round(data.loc[mask_c8m,  'Conc.'].sum(), 2)
    c16plus    = round(data.loc[mask_c16p, 'Conc.'].sum(), 2)
    filt816    = data[mask_c816]
    c816_total = round(filt816['Conc.'].sum(), 2)
    sums_816   = filt816.groupby('Category', dropna=True)['Conc.'].sum()

    iso_val   = sums_816.get('iso-alkane', 0.0)
    n_val     = sums_816.get('n-alkane',   0.0)
    i_n_ratio = round(iso_val / n_val, 2) if n_val > 0 else 0.0

    # Step 6: 找或新增 Summary 表格
    summary_header_row = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == 'C8-':
                summary_header_row = cell.row
                break
        if summary_header_row:
            break

    # 若原始檔沒有 Summary 表格，在 Peak Table header 上面兩列新增
    if summary_header_row is None:
        summary_header_row = max(1, header_row_idx - 1)  # 在 header 上方插入
        ws.insert_rows(summary_header_row, amount=2)
        # 重新計算 DATA_START_ROW（插入了 2 列）
        DATA_START_ROW += 2

        # 寫入 Summary header
        summary_labels = ['C8-', 'C8-16', 'C16+', 'i/n',
                          'alkene', 'aromatic', 'cyclic', 'iso', 'n-alkane', 'others', 'poly']
        header_colors  = ['FFD966', 'FFD966', 'FFD966', 'FFD966',
                          'ADD8E6', '92D050', 'FFD700', 'FF6600', 'FF8C00', 'FA8072', 'C5B0D5']
        for col_offset, (label, color) in enumerate(zip(summary_labels, header_colors)):
            cell = ws.cell(row=summary_header_row, column=col_offset + 1)
            cell.value     = label
            cell.font      = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.fill      = PatternFill(start_color=color, end_color=color, fill_type='solid')

    SUMMARY_VALUES_ROW = summary_header_row + 1
    col_map = {}
    for cell in ws[summary_header_row]:
        if cell.value:
            col_map[str(cell.value).strip()] = cell.column

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

    # Step 7: 驗證
    grand_total    = c8minus + c816_total + c16plus
    all_conc_sum   = round(data['Conc.'].sum(), 2)
    error          = abs(grand_total - all_conc_sum)
    passed         = error < 0.01
    others_816     = filt816[filt816['Category'] == 'others'][['Ret.Time', 'Name', 'Conc.']].to_dict('records')

    # Step 8: 把驗證報告寫入 xlsx（Summary 表格下方空一列後新增）
    report_start_row = SUMMARY_VALUES_ROW + 2

    # 標題列
    title_cell = ws.cell(row=report_start_row, column=1)
    title_cell.value = 'Validation Report'
    title_cell.font  = Font(bold=True, size=11)
    title_cell.fill  = PatternFill(start_color='2B6CB0', end_color='2B6CB0', fill_type='solid')
    title_cell.font  = Font(bold=True, color='FFFFFF', size=11)

    # 各項數值
    report_rows = [
        ('C8-  (conc%)',             c8minus),
        ('C8-16  (conc%)',           c816_total),
        ('C16+  (conc%)',            c16plus),
        ('  Aromatic',               round(sums_816.get('Aromatic',     0.0), 2)),
        ('  Polyaromatic',           round(sums_816.get('Polyaromatic', 0.0), 2)),
        ('  n-alkane',               round(sums_816.get('n-alkane',     0.0), 2)),
        ('  iso-alkane',             round(sums_816.get('iso-alkane',   0.0), 2)),
        ('  cyclic',                 round(sums_816.get('cyclic',       0.0), 2)),
        ('  Alkene',                 round(sums_816.get('Alkene',       0.0), 2)),
        ('  others',                 round(sums_816.get('others',       0.0), 2)),
        ('Total check (sum)',         grand_total),
        ('All Conc. sum',            all_conc_sum),
        ('Δ error',                  round(error, 4)),
        ('PASS / FAIL',              'PASS ✓' if passed else f'FAIL (Δ={round(error,4)})'),
        ('Auto-classified rows',     auto_classified_count),
    ]

    for offset, (label, value) in enumerate(report_rows):
        r = report_start_row + 1 + offset
        ws.cell(row=r, column=1).value = label
        ws.cell(row=r, column=2).value = value
        ws.cell(row=r, column=1).font  = Font(size=10)
        ws.cell(row=r, column=2).font  = Font(bold=True, size=10)
        if label == 'PASS / FAIL':
            color = 'C6EFCE' if passed else 'FFC7CE'
            ws.cell(row=r, column=2).fill = PatternFill(
                start_color=color, end_color=color, fill_type='solid')

    # 若有 others 清單就列出來
    if others_816:
        others_start = report_start_row + 1 + len(report_rows) + 1
        ws.cell(row=others_start, column=1).value = 'Unclassified (others) in C8-16'
        ws.cell(row=others_start, column=1).font  = Font(bold=True, color='FF0000')
        for i, rec in enumerate(others_816):
            r = others_start + 1 + i
            ws.cell(row=r, column=1).value = rec.get('Name', '')
            ws.cell(row=r, column=2).value = rec.get('Ret.Time', '')
            ws.cell(row=r, column=3).value = rec.get('Conc.', '')

    # 儲存
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    validation_report = {
        'c8minus':    c8minus,
        'c816_total': c816_total,
        'c16plus':    c16plus,
        'grand_total':    grand_total,
        'all_conc_sum':   all_conc_sum,
        'error':          round(error, 4),
        'passed':         passed,
        'auto_classified': auto_classified_count,
        'others_816':     others_816,
        'sums_816': {k: round(v, 2) for k, v in sums_816.items()},
    }

    return out.read(), validation_report


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/classify', methods=['POST'])
def classify():
    if 'file' not in request.files:
        return jsonify({'error': '沒有收到檔案'}), 400

    f = request.files['file']
    if not f.filename.endswith('.xlsx'):
        return jsonify({'error': '只接受 .xlsx 格式'}), 400

    file_bytes = f.read()
    original_filename = secure_filename(f.filename)

    try:
        result_bytes, report = process_gcms_file(file_bytes, original_filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    stem = original_filename.rsplit('.', 1)[0]
    out_filename = f"{stem}_分類.xlsx"

    response = send_file(
        io.BytesIO(result_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=out_filename,
    )
    response.headers['X-C8minus']    = str(report['c8minus'])
    response.headers['X-C816total']  = str(report['c816_total'])
    response.headers['X-C16plus']    = str(report['c16plus'])
    response.headers['X-GrandTotal'] = str(report['grand_total'])
    response.headers['X-Passed']     = str(report['passed'])
    response.headers['X-AutoClass']  = str(report['auto_classified'])
    s816 = report['sums_816']
    response.headers['X-Aromatic']     = str(round(s816.get('Aromatic',     0.0), 2))
    response.head