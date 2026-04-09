"""
台達 (Delta) Forecast 合併處理器
================================
將 3 個不同格式的 Buyer Forecast 檔案合併為統一匯總格式。

Buyer 格式:
  Ketwadee (PSB5): MRP sheet, 3 rows/part (Demand/Supply/Net)
  Kanyanat (PSB7): Sheet1, 4 rows/part (A-Demand/B-Supply/D-Net/E-REMARK)
  Weeraya  (PSB7): Sheet1, 5 rows/part (Demand/Firmed/ForecastConf/NetDemand/Remark)

匯總格式: 每料號 3 行 (Demand/Supply/Balance), Balance 用 Excel 公式
"""

import openpyxl
import os
from datetime import datetime
from copy import copy
from openpyxl.styles import (Font, Alignment, PatternFill, Border, Side)
from openpyxl.styles.colors import Color


# ---------------------------------------------------------------------------
# Buyer 自動偵測
# ---------------------------------------------------------------------------

def detect_buyer(filepath):
    """
    自動偵測 Forecast 檔案屬於哪個 Buyer。
    Returns: 'Ketwadee' | 'Kanyanat' | 'Weeraya' | None
    """
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        sheet_names = wb.sheetnames

        # Ketwadee: 有 MRP sheet
        if 'MRP' in sheet_names:
            wb.close()
            return 'Ketwadee'

        # Kanyanat / Weeraya 都用 Sheet1
        if 'Sheet1' not in sheet_names:
            wb.close()
            return None

        ws = wb['Sheet1']

        # Kanyanat: col X (24) 有 'A-Demand'
        for row in ws.iter_rows(min_row=2, max_row=min(20, ws.max_row or 20),
                                min_col=24, max_col=24, values_only=True):
            if row[0] and 'A-Demand' in str(row[0]):
                wb.close()
                return 'Kanyanat'

        # Weeraya: col L (12) 有 'Demand'
        for row in ws.iter_rows(min_row=2, max_row=min(20, ws.max_row or 20),
                                min_col=12, max_col=12, values_only=True):
            if row[0] and str(row[0]).strip() == 'Demand':
                wb.close()
                return 'Weeraya'

        wb.close()
        return None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 日期欄位工具
# ---------------------------------------------------------------------------

MONTH_NAMES = ('JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
               'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC')

# 每個 Buyer header 中日期欄位的起始 column (1-based)
_BUYER_DATE_START_COL = {
    'Ketwadee': 16,   # col P onwards (MRP sheet)
    'Kanyanat': 25,   # col Y onwards (Sheet1)
    'Weeraya':  14,   # col N onwards (Sheet1)
}
_BUYER_SHEET = {
    'Ketwadee': 'MRP',
    'Kanyanat': 'Sheet1',
    'Weeraya':  'Sheet1',
}


def read_date_cols_from_template(template_path):
    """
    從匯總格式模板的 header row 動態讀取日期欄位。(備用，正式流程用 buyer 檔案提取)
    Returns: list of normalized date column names (e.g. ['PASSDUE','20260406',...,'MAR'])
    """
    wb = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    date_cols = []
    for cell in ws[1]:
        if cell.column >= 10 and cell.value is not None:  # J (col 10) onwards
            date_cols.append(_normalize_date_header(cell.value))
    wb.close()
    return [d for d in date_cols if d is not None]


def _normalize_date_header(val):
    """將不同源頭的日期格式統一為匯總格式的 key"""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime('%Y%m%d')
    s = str(val).strip()
    if 'PAST' in s.upper() or 'PASSDUE' in s.upper():
        return 'PASSDUE'
    if 'OVER DUE' in s.upper():
        return 'PASSDUE'
    if s.isdigit() and len(s) == 8:
        return s
    # "2026-JUL" → "JUL" (不限定年份，自動適用任何年度)
    if '-' in s and len(s) >= 5:
        parts = s.split('-')
        if len(parts) == 2 and len(parts[1]) == 3:
            month = parts[1].upper()
            if month in MONTH_NAMES:
                return month
    # 直接就是月份名 (JUL, AUG, ...)
    if s.upper() in MONTH_NAMES:
        return s.upper()
    return None


def _sort_date_cols(dates):
    """
    排序日期欄位: PASSDUE → 週日期(YYYYMMDD) → 月份(從最後週日期的月份開始)
    """
    passdue = [d for d in dates if d == 'PASSDUE']
    weekly = sorted([d for d in dates if d.isdigit() and len(d) == 8])
    monthly = [d for d in dates if d in MONTH_NAMES]

    if weekly and monthly:
        # 根據最後一個週日期的月份決定月份排序起點
        last_month_idx = int(weekly[-1][4:6]) - 1  # 0-based
        rotated = list(MONTH_NAMES[last_month_idx:]) + list(MONTH_NAMES[:last_month_idx])
        monthly.sort(key=lambda m: rotated.index(m))

    return passdue + weekly + monthly


def extract_dates_from_buyer_files(buyer_files):
    """
    從 Buyer 原始檔案的 header 動態提取所有日期欄位。

    Args:
        buyer_files: dict {buyer_name: filepath} e.g. {'Ketwadee': 'path/to/file.xlsx'}

    Returns:
        tuple(date_cols, per_buyer_dates, warnings)
        - date_cols: list — 排序後的統一日期欄位 (聯集)
        - per_buyer_dates: dict — 每個 Buyer 的日期集合 (供比對)
        - warnings: list — 日期不一致的警告訊息
    """
    per_buyer_dates = {}

    for buyer_name, filepath in buyer_files.items():
        sheet_name = _BUYER_SHEET.get(buyer_name)
        start_col = _BUYER_DATE_START_COL.get(buyer_name, 10)

        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active

        buyer_dates = set()
        for cell in ws[1]:
            if cell.column >= start_col and cell.value is not None:
                norm = _normalize_date_header(cell.value)
                if norm is not None and norm not in buyer_dates:
                    buyer_dates.add(norm)  # 同名只取第一個 (避免 2026-JUL / 2027-JUL 重複)
        wb.close()

        per_buyer_dates[buyer_name] = buyer_dates
        print(f"  {buyer_name}: 偵測到 {len(buyer_dates)} 個日期欄位")

    # 聯集 = 所有 buyer 的日期
    all_dates = set()
    for dates in per_buyer_dates.values():
        all_dates |= dates

    # 比對差異
    warnings = []
    buyers = list(per_buyer_dates.keys())
    for i, b1 in enumerate(buyers):
        for b2 in buyers[i + 1:]:
            only_in_b1 = per_buyer_dates[b1] - per_buyer_dates[b2]
            only_in_b2 = per_buyer_dates[b2] - per_buyer_dates[b1]
            if only_in_b1:
                msg = f'{b1} 有但 {b2} 沒有的日期: {sorted(only_in_b1)}'
                warnings.append(msg)
                print(f"  ⚠️ {msg}")
            if only_in_b2:
                msg = f'{b2} 有但 {b1} 沒有的日期: {sorted(only_in_b2)}'
                warnings.append(msg)
                print(f"  ⚠️ {msg}")

    if not warnings:
        print(f"  ✅ 3 個 Buyer 日期完全一致 ({len(all_dates)} 個日期)")

    date_cols = _sort_date_cols(all_dates)
    return date_cols, per_buyer_dates, warnings


# ---------------------------------------------------------------------------
# Buyer 讀取器
# ---------------------------------------------------------------------------

def _build_date_col_map(ws, start_col, date_cols):
    """建立 column → date_key 的映射，避免同名碰撞 (如 2026-JUL 和 2027-JUL)"""
    date_col_map = {}
    seen_keys = set()
    for cell in ws[1]:
        if cell.value is not None and cell.column >= start_col:
            norm = _normalize_date_header(cell.value)
            if norm in date_cols and norm not in seen_keys:
                date_col_map[cell.column] = norm
                seen_keys.add(norm)
    return date_col_map


def _to_partno(val):
    """將 PARTNO 轉為數字（如果可能），與客戶格式一致"""
    if val is None:
        return ''
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return str(val)


def extract_plant_codes_from_regions(regions):
    """
    從 mapping 表的 region 欄位提取 PLANT 代碼。
    例: ['PSB5 泰國', 'IPC1 東莞'] → ['PSB5', 'IPC1']
    """
    codes = []
    for r in regions:
        if not r:
            continue
        parts = str(r).split()
        if parts:
            codes.append(parts[0])
    return codes


def match_plants_in_filename(filepath, plant_codes):
    """
    從檔名中搜尋 PLANT 代碼 (case-insensitive)。
    回傳檔名中出現的所有 PLANT 代碼，依長度降冪排序 (長的優先)。

    Args:
        filepath: 檔案路徑
        plant_codes: 有效的 PLANT 代碼清單 (如 ['PSB5', 'PSB7', 'IPC1'])

    Returns:
        list of matched codes, e.g. ['PSB5'] or ['IAI1', 'UPI2', 'DFI1']
    """
    if not plant_codes:
        return []
    fn = os.path.splitext(os.path.basename(filepath))[0].upper()
    matched = [code for code in plant_codes if code.upper() in fn]
    matched.sort(key=len, reverse=True)
    return matched


def _read_ketwadee(filepath, date_cols, buyer_label=None, plant_code=None):
    """讀取 PSB5 Ketwadee: MRP sheet, 3 rows/part"""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['MRP']

    date_col_map = _build_date_col_map(ws, 16, date_cols)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        filter_val = row[14].value if len(row) > 14 else None

        if filter_val == 'Demand':
            part_no = row[2].value
            vendor_part = row[3].value
            stock = row[7].value or 0

            demand = {}
            for col_idx, date_key in date_col_map.items():
                if col_idx - 1 < len(row):
                    v = row[col_idx - 1].value
                    demand[date_key] = v if v is not None else 0
                else:
                    demand[date_key] = 0

            supply = {}
            if i + 1 < len(rows):
                supply_row = rows[i + 1]
                for col_idx, date_key in date_col_map.items():
                    if col_idx - 1 < len(supply_row):
                        v = supply_row[col_idx - 1].value
                        supply[date_key] = v if v is not None else 0
                    else:
                        supply[date_key] = 0

            results.append({
                'buyer': buyer_label or 'Ketwadee', 'plant': plant_code or 'PSB5',
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock, 'on_way': None,
                'demand': demand, 'supply': supply,
            })
            i += 3
        else:
            i += 1

    wb.close()
    return results


def _read_kanyanat(filepath, date_cols, buyer_label=None, plant_code=None):
    """讀取 PSB7 Kanyanat: Sheet1, 4 rows/part"""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 25, date_cols)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        type_val = row[23].value if len(row) > 23 else None

        if type_val == 'A-Demand':
            part_no = row[4].value
            vendor_part = row[5].value

            demand = {}
            for col_idx, date_key in date_col_map.items():
                if col_idx - 1 < len(row):
                    v = row[col_idx - 1].value
                    demand[date_key] = v if v is not None else 0
                else:
                    demand[date_key] = 0

            supply = {}
            if i + 1 < len(rows):
                supply_row = rows[i + 1]
                for col_idx, date_key in date_col_map.items():
                    if col_idx - 1 < len(supply_row):
                        v = supply_row[col_idx - 1].value
                        supply[date_key] = v if v is not None else 0
                    else:
                        supply[date_key] = 0

            results.append({
                'buyer': buyer_label or 'Kanyanat', 'plant': plant_code or 'PSB7',
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': None, 'on_way': None,
                'demand': demand, 'supply': supply,
            })
            i += 4
        else:
            i += 1

    wb.close()
    return results


def _read_weeraya(filepath, date_cols, buyer_label=None, plant_code=None):
    """讀取 PSB7 Weeraya: Sheet1, 5 rows/part"""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 14, date_cols)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        type_val = row[11].value if len(row) > 11 else None

        if type_val == 'Demand':
            part_no = row[3].value
            vendor_part = row[4].value
            stock = row[12].value or 0

            def read_date_values(r):
                data = {}
                for col_idx, date_key in date_col_map.items():
                    if col_idx - 1 < len(r):
                        v = r[col_idx - 1].value
                        data[date_key] = v if v is not None else 0
                    else:
                        data[date_key] = 0
                return data

            demand = read_date_values(row)
            supply = read_date_values(rows[i + 2]) if i + 2 < len(rows) else {}
            balance_data = read_date_values(rows[i + 3]) if i + 3 < len(rows) else {}

            results.append({
                'buyer': buyer_label or 'Weeraya', 'plant': plant_code or 'PSB7',
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock, 'on_way': None,
                'demand': demand, 'supply': supply,
                'balance_override': balance_data,
            })
            i += 5
        else:
            i += 1

    wb.close()
    return results


# ---------------------------------------------------------------------------
# Excel 產出
# ---------------------------------------------------------------------------

def _generate_consolidated_excel(all_source, output_path, reference_path,
                                 date_cols, erp_mapping=None):
    """
    產出匯總格式 Excel, 格式與原始模板完全一致。
    Balance 行使用 Excel 公式。
    C/D 欄位 (ERP 客戶簡稱、ERP 送貨地點) 留空，由第三步驟 forecast 處理時
    透過 customer_mappings 表自動填入，consolidation 階段不寫入。
    """
    # erp_mapping 已棄用，保留參數做相容性

    wb_ref = openpyxl.load_workbook(reference_path)
    ws_ref = wb_ref[wb_ref.sheetnames[0]]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '工作表1'

    # 複製固定表頭 A~I (col 1~9) 從模板
    for col in range(1, 10):
        src = ws_ref.cell(row=1, column=col)
        dst = ws.cell(row=1, column=col, value=src.value)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.border = copy(src.border)

    # 動態寫入日期表頭 J onwards (從 date_cols 產生，不依賴模板)
    # 取得模板中 J 欄的樣式作為日期表頭的參考樣式
    ref_date_cell = ws_ref.cell(row=1, column=10)
    for col_offset, date_key in enumerate(date_cols):
        col_num = 10 + col_offset
        dst = ws.cell(row=1, column=col_num, value=date_key)
        dst.font = copy(ref_date_cell.font)
        dst.fill = copy(ref_date_cell.fill)
        dst.alignment = copy(ref_date_cell.alignment)
        dst.number_format = ref_date_cell.number_format
        dst.border = copy(ref_date_cell.border)

    # 複製固定欄位 A~I 的欄寬
    from openpyxl.utils import get_column_letter
    for col in range(1, 10):
        cl = get_column_letter(col)
        if cl in ws_ref.column_dimensions:
            ws.column_dimensions[cl].width = ws_ref.column_dimensions[cl].width
    # 日期欄位統一寬度 (參考模板 J 欄)
    ref_j_width = ws_ref.column_dimensions.get('J')
    date_col_width = ref_j_width.width if ref_j_width else 11
    for col_offset in range(len(date_cols)):
        cl = get_column_letter(10 + col_offset)
        ws.column_dimensions[cl].width = date_col_width

    # 樣式定義
    font_mingliu_12 = Font(name='新細明體', size=12)
    font_arial_9 = Font(name='Arial', size=9)
    font_arial_10 = Font(name='Arial', size=10)
    nf_data = '#,##0_);[Red]\\(#,##0\\)'
    nf_partno = '000'
    nf_stock = '#,##0'
    align_center = Alignment(vertical='center')
    align_top_left = Alignment(horizontal='left', vertical='top')
    fill_cd = PatternFill(patternType='solid',
                          fgColor=Color(theme=4, tint=0.7999816888943144))
    fill_i = PatternFill(patternType='solid',
                         fgColor=Color(theme=0, tint=0.0))
    fill_supply = PatternFill(patternType='solid',
                              fgColor=Color(theme=9, tint=0.7999816888943144))
    fill_balance = PatternFill(patternType='solid',
                               fgColor=Color(theme=7, tint=0.7999816888943144))
    thin_side = Side(style='thin')
    medium_side = Side(style='medium')
    border_normal = Border(left=thin_side, right=thin_side,
                           top=thin_side, bottom=thin_side)
    border_balance = Border(left=thin_side, right=thin_side,
                            top=thin_side, bottom=medium_side)

    def apply_row_style(row_num, date_type):
        is_balance = (date_type == 'Balance')
        bdr = border_balance if is_balance else border_normal

        for col in [1, 2]:
            c = ws.cell(row=row_num, column=col)
            c.alignment = align_center
            c.border = bdr
        for col in [3, 4]:
            c = ws.cell(row=row_num, column=col)
            c.alignment = align_center
            c.border = bdr
            c.fill = fill_cd
        for col in [5, 6]:
            c = ws.cell(row=row_num, column=col)
            c.alignment = align_top_left
            c.border = bdr
        c = ws.cell(row=row_num, column=7)
        c.font = font_arial_9
        c.alignment = align_top_left
        c.border = bdr
        c = ws.cell(row=row_num, column=8)
        c.font = font_mingliu_12
        c.alignment = align_center
        c.border = bdr
        c = ws.cell(row=row_num, column=9)
        c.alignment = align_top_left
        c.border = bdr
        c.fill = fill_i

        data_fill = fill_supply if date_type == 'Supply' else (
            fill_balance if date_type == 'Balance' else None)
        for col in range(10, 10 + len(date_cols)):
            c = ws.cell(row=row_num, column=col)
            c.alignment = align_center
            c.border = bdr
            if data_fill:
                c.fill = data_fill

    # 寫入資料
    row_num = 2
    for item in all_source:
        buyer = item['buyer']
        plant = item['plant']
        part_no = item['part_no']
        vendor_part = item['vendor_part']
        stock = item['stock']
        # ERP 客戶簡稱、ERP 送貨地點 留空 (第三步驟由 mapping 表填入)
        erp_name, erp_location = '', ''

        # --- Demand ---
        demand_row = row_num
        ws.cell(row=row_num, column=1, value=buyer).font = font_mingliu_12
        ws.cell(row=row_num, column=2, value=plant).font = font_mingliu_12
        ws.cell(row=row_num, column=3, value=erp_name).font = font_mingliu_12
        ws.cell(row=row_num, column=4, value=erp_location).font = font_mingliu_12
        c = ws.cell(row=row_num, column=5, value=part_no)
        c.font = font_arial_9
        c.number_format = nf_partno
        ws.cell(row=row_num, column=6, value=vendor_part).font = font_arial_9
        if stock is not None:
            c = ws.cell(row=row_num, column=7, value=stock)
            c.font = font_arial_9
            c.number_format = nf_stock
        ws.cell(row=row_num, column=9, value='Demand').font = font_arial_10
        for col_offset, date_key in enumerate(date_cols):
            v = item['demand'].get(date_key, 0) or 0
            c = ws.cell(row=row_num, column=10 + col_offset, value=v)
            c.font = font_mingliu_12
            c.number_format = nf_data
        apply_row_style(row_num, 'Demand')
        row_num += 1

        # --- Supply ---
        supply_row = row_num
        ws.cell(row=row_num, column=1, value=buyer).font = font_mingliu_12
        ws.cell(row=row_num, column=2, value=plant).font = font_mingliu_12
        ws.cell(row=row_num, column=3, value=erp_name).font = font_mingliu_12
        ws.cell(row=row_num, column=4, value=erp_location).font = font_mingliu_12
        c = ws.cell(row=row_num, column=5, value=part_no)
        c.font = font_arial_9
        c.number_format = nf_partno
        ws.cell(row=row_num, column=6, value=vendor_part).font = font_arial_9
        ws.cell(row=row_num, column=9, value='Supply').font = font_arial_10
        for col_offset, date_key in enumerate(date_cols):
            v = item['supply'].get(date_key, 0) or 0
            c = ws.cell(row=row_num, column=10 + col_offset,
                        value=v if v != 0 else None)
            c.font = font_mingliu_12
            c.number_format = nf_data
        apply_row_style(row_num, 'Supply')
        row_num += 1

        # --- Balance (公式) ---
        balance_row = row_num
        ws.cell(row=row_num, column=1, value=buyer).font = font_mingliu_12
        ws.cell(row=row_num, column=2, value=plant).font = font_mingliu_12
        ws.cell(row=row_num, column=3, value=erp_name).font = font_mingliu_12
        ws.cell(row=row_num, column=4, value=erp_location).font = font_mingliu_12
        c = ws.cell(row=row_num, column=5, value=part_no)
        c.font = font_arial_9
        c.number_format = nf_partno
        ws.cell(row=row_num, column=6, value=vendor_part).font = font_arial_9
        ws.cell(row=row_num, column=9, value='Balance').font = font_arial_10
        for col_offset, date_key in enumerate(date_cols):
            col_num = 10 + col_offset
            col_letter = openpyxl.utils.get_column_letter(col_num)
            if col_offset == 0:
                formula = f'=G{demand_row}+H{demand_row}-{col_letter}{demand_row}'
            elif col_offset == 1:
                prev_col = openpyxl.utils.get_column_letter(col_num - 1)
                formula = (f'={prev_col}{supply_row}+{prev_col}{balance_row}'
                           f'-{col_letter}{demand_row}')
            else:
                prev_col = openpyxl.utils.get_column_letter(col_num - 1)
                formula = (f'={prev_col}{balance_row}+{prev_col}{supply_row}'
                           f'-{col_letter}{demand_row}')
            c = ws.cell(row=balance_row, column=col_num, value=formula)
            c.font = font_mingliu_12
            c.number_format = nf_data
        apply_row_style(row_num, 'Balance')
        row_num += 1

    # 凍結窗格、自動篩選
    ws.freeze_panes = 'J2'
    last_data_col = 9 + len(date_cols)  # A~I (9) + 日期欄位數
    last_col = openpyxl.utils.get_column_letter(last_data_col)
    ws.auto_filter.ref = f'A1:{last_col}{row_num - 1}'

    # 條件式格式
    for cf in ws_ref.conditional_formatting:
        for rule in cf.rules:
            ws.conditional_formatting.add(str(cf.cells), rule)

    # 複製工作表5 (如果存在)
    if '工作表5' in wb_ref.sheetnames:
        ws5_ref = wb_ref['工作表5']
        ws5 = wb.create_sheet('工作表5')
        for row in ws5_ref.iter_rows(min_row=1, max_row=ws5_ref.max_row,
                                      max_col=ws5_ref.max_column, values_only=False):
            for cell in row:
                dst = ws5.cell(row=cell.row, column=cell.column, value=cell.value)
                dst.font = copy(cell.font)
                dst.fill = copy(cell.fill)
                dst.alignment = copy(cell.alignment)
                dst.number_format = cell.number_format
                dst.border = copy(cell.border)
        for cl, dim in ws5_ref.column_dimensions.items():
            ws5.column_dimensions[cl].width = dim.width

    wb_ref.close()
    wb.save(output_path)
    return len(all_source)  # 料號數


# ---------------------------------------------------------------------------
# 主要入口
# ---------------------------------------------------------------------------

BUYER_READERS = {
    'Ketwadee': _read_ketwadee,
    'Kanyanat': _read_kanyanat,
    'Weeraya': _read_weeraya,
}


def consolidate(forecast_files, reference_template, output_path,
                erp_mapping=None, plant_codes=None):
    """
    合併多個 Delta Forecast 檔案為匯總格式 Excel。

    Args:
        forecast_files: list of file paths (3 個 Buyer 檔案)
        reference_template: 匯總格式模板路徑 (用於取得表頭格式，不用於日期)
        output_path: 輸出檔案路徑
        erp_mapping: dict {Plant: (ERP客戶簡稱, ERP送貨地點)}, 已棄用
        plant_codes: list of valid PLANT codes, 用於從檔名比對 PLANT。
                     若為 None 則 fallback 到舊的寫死值 (PSB5/PSB7)。
                     建議由 customer_mappings 的 region 欄位提取。

    Returns:
        dict with keys: success, part_count, buyer_stats, date_warnings, message
    """
    # 1. 偵測每個檔案的 Buyer
    buyer_files = {}
    for fp in forecast_files:
        buyer = detect_buyer(fp)
        if buyer is None:
            return {
                'success': False, 'part_count': 0,
                'message': f'無法識別檔案格式: {os.path.basename(fp)}'
            }
        if buyer in buyer_files:
            return {
                'success': False, 'part_count': 0,
                'message': f'重複的 Buyer: {buyer} (每個 Buyer 只能上傳一個檔案)'
            }
        buyer_files[buyer] = fp

    print(f"Delta 合併: 偵測到 {len(buyer_files)} 個 Buyer: {list(buyer_files.keys())}")

    # 2. 從 Buyer 原始檔案動態提取日期欄位（不依賴模板）
    date_cols, per_buyer_dates, date_warnings = extract_dates_from_buyer_files(buyer_files)
    if not date_cols:
        return {
            'success': False, 'part_count': 0,
            'message': '無法從 Buyer 檔案提取日期欄位'
        }

    print(f"  統一日期欄位: {len(date_cols)} 個 ({date_cols[0]} ~ {date_cols[-1]})")

    # 3. 讀取各 Buyer 資料
    # Buyer 欄位暫時使用檔案名稱 (不含副檔名)，等客戶確認代碼→名稱對照後再調整邏輯
    # PLANT 欄位從 mapping 表的 region 提取英文代碼 (如 'PSB5 泰國' → 'PSB5')，
    # 再從檔名中比對出現的代碼 (大小寫無關)。
    all_source = []
    buyer_stats = {}
    for buyer_name in ['Ketwadee', 'Kanyanat', 'Weeraya']:
        if buyer_name in buyer_files:
            reader = BUYER_READERS[buyer_name]
            fp = buyer_files[buyer_name]
            buyer_label = os.path.splitext(os.path.basename(fp))[0]

            # 從檔名比對 PLANT 代碼
            plant_code = None
            if plant_codes:
                matched = match_plants_in_filename(fp, plant_codes)
                if matched:
                    plant_code = matched[0]  # 單一 PLANT 檔案取第一個
                    if len(matched) > 1:
                        print(f"  ⚠️ {buyer_label}: 檔名中有多個 PLANT {matched}, 取 {plant_code}")

            data = reader(fp, date_cols,
                          buyer_label=buyer_label, plant_code=plant_code)
            buyer_stats[buyer_name] = len(data)
            all_source.extend(data)
            plant_display = plant_code or '(寫死預設值)'
            print(f"  {buyer_name} ({buyer_label}): {len(data)} 個料號, PLANT={plant_display}")

    if not all_source:
        return {
            'success': False, 'part_count': 0,
            'message': '未讀取到任何料號資料'
        }

    # 4. 產出匯總格式 Excel（日期表頭由 date_cols 動態產生）
    part_count = _generate_consolidated_excel(
        all_source, output_path, reference_template, date_cols, erp_mapping
    )

    print(f"  合併完成: {part_count} 個料號 → {os.path.basename(output_path)}")

    result = {
        'success': True,
        'part_count': part_count,
        'buyer_stats': buyer_stats,
        'message': f'成功合併 {part_count} 個料號'
    }
    if date_warnings:
        result['date_warnings'] = date_warnings
    return result
