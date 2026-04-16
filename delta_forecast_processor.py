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
from datetime import datetime, timedelta
from copy import copy
from openpyxl.styles import (Font, Alignment, PatternFill, Border, Side)
from openpyxl.styles.colors import Color


# ---------------------------------------------------------------------------
# 格式常數
# ---------------------------------------------------------------------------

FORMAT_KETWADEE = 'ketwadee'              # MRP sheet, 3 rows/part
FORMAT_KANYANAT = 'kanyanat'              # Sheet1, 4 rows/part (col 24=TYPE)
FORMAT_WEERAYA = 'weeraya'                # Sheet1, 5 rows/part (col 12=TYPE)
FORMAT_INDIA_IAI1 = 'india_iai1'          # PAN JIT, 3 rows/part, 多PLANT
FORMAT_PSW1_CEW1 = 'psw1_cew1'            # Sheet1, 5 rows/part (col 12=Status), 多PLANT
FORMAT_MWC1IPC1 = 'mwc1ipc1'              # Sheet1, 4 rows/part (col 6=REQUEST ITEM), 多PLANT
FORMAT_NBQ1 = 'nbq1'                      # PAN JIT, 1 row/part, 單PLANT檔名
FORMAT_SVC1PWC1_DIODE_MOS = 'svc1pwc1_diode_mos'  # Diode+MOS, 1 row/part, 多PLANT
FORMAT_PSBG = 'psbg'                              # Sheet1, 3 rows/part (col 15=Filter), 單PLANT
FORMAT_EIBG_EISBG = 'eibg_eisbg'                  # Sheet1, col1=ITEM, flat (Demand only), 單PLANT
FORMAT_FMBG = 'fmbg'                              # Sheet1, col12=REQUEST ITEM, 3 rows/part (A-Demand/B-CFM/C-Bal), 多PLANT
FORMAT_IABG = 'iabg'                              # Sheet1, col9=SHIP, flat (Demand only), 多PLANT
FORMAT_ICTBG_NTL7 = 'ictbg_ntl7'                  # Sheet1, col10=REQUEST ITEM, 4 rows/part (GROSS REQTS/...), 多PLANT
FORMAT_ICTBG_PSB9_MRP = 'ictbg_psb9_mrp'          # PSB9_MRP* sheet, col14=Type, 4 rows/part (DEMAND/SUPPLY/NET/Remark), 多PLANT
FORMAT_ICTBG_PSB9_SIRIRAHT = 'ictbg_psb9_siriraht'  # Sheet1, col15=REQUEST ITEM (1Demand/2Supply/3Balance), 3 rows/part, 多PLANT

FORMAT_LABELS = {
    FORMAT_KETWADEE: 'Ketwadee (PSB5)',
    FORMAT_KANYANAT: 'Kanyanat (PSB7)',
    FORMAT_WEERAYA:  'Weeraya (PSB7)',
    FORMAT_INDIA_IAI1: 'India IAI1/UPI2/DFI1',
    FORMAT_PSW1_CEW1: 'PSW1+CEW1',
    FORMAT_MWC1IPC1:  'MWC1+IPC1',
    FORMAT_NBQ1:      'NBQ1',
    FORMAT_SVC1PWC1_DIODE_MOS: 'SVC1+PWC1 (Diode&MOS)',
    FORMAT_PSBG:     'PSBG (PSB5 PANJIT)',
    FORMAT_EIBG_EISBG: 'EIBG/EISBG (UPW1)',
    FORMAT_FMBG:       'FMBG (TPC5/EMN3)',
    FORMAT_IABG:       'IABG (IMW1)',
    FORMAT_ICTBG_NTL7: 'ICTBG (NTL7)',
    FORMAT_ICTBG_PSB9_MRP:      'ICTBG PSB9 Kaewarin',
    FORMAT_ICTBG_PSB9_SIRIRAHT: 'ICTBG PSB9 Siriraht',
}

# 單 PLANT 檔案 (PLANT 從檔名比對)
SINGLE_PLANT_FORMATS = {
    FORMAT_KETWADEE, FORMAT_KANYANAT, FORMAT_WEERAYA, FORMAT_NBQ1, FORMAT_PSBG,
    FORMAT_EIBG_EISBG,
}
# 多 PLANT 檔案 (PLANT 從檔案每列讀)
MULTI_PLANT_FORMATS = {
    FORMAT_INDIA_IAI1, FORMAT_PSW1_CEW1, FORMAT_MWC1IPC1, FORMAT_SVC1PWC1_DIODE_MOS,
    FORMAT_FMBG, FORMAT_IABG, FORMAT_ICTBG_NTL7, FORMAT_ICTBG_PSB9_MRP,
    FORMAT_ICTBG_PSB9_SIRIRAHT,
}


# ---------------------------------------------------------------------------
# 格式自動偵測
# ---------------------------------------------------------------------------

def _cell_str(ws, row, col):
    """安全取 cell 字串值"""
    try:
        v = ws.cell(row=row, column=col).value
        return str(v).strip() if v is not None else ''
    except Exception:
        return ''


def detect_format(filepath):
    """
    自動偵測 Delta Forecast 檔案屬於哪種格式。
    Returns: FORMAT_* 常數之一，無法識別則回傳 None
    """
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        sheets = wb.sheetnames

        # === 最明確: 兩個 sheet 組合 ===
        if 'Diode' in sheets and 'MOS' in sheets:
            wb.close()
            return FORMAT_SVC1PWC1_DIODE_MOS

        # === 唯一 sheet 名稱: MRP → Ketwadee ===
        if 'MRP' in sheets:
            wb.close()
            return FORMAT_KETWADEE

        # === PSB9_MRP* sheet → ICTBG PSB9 Kaewarin ===
        for s in sheets:
            if s.startswith('PSB9_MRP'):
                wb.close()
                return FORMAT_ICTBG_PSB9_MRP

        # === PAN JIT sheet: 分辨 India IAI1 vs NBQ1 ===
        if 'PAN JIT' in sheets:
            ws = wb['PAN JIT']
            h1 = _cell_str(ws, 1, 1)
            h3 = _cell_str(ws, 1, 3)
            h13 = _cell_str(ws, 1, 13)
            # India IAI1: col 3 = PLANT, col 13 = Request
            if h3.upper() == 'PLANT' and h13.lower() == 'request':
                wb.close()
                return FORMAT_INDIA_IAI1
            # NBQ1: col 1 = PARTNO, no PLANT/Request
            if h1.upper() == 'PARTNO':
                wb.close()
                return FORMAT_NBQ1

        # === Sheet1: 多種格式 ===
        if 'Sheet1' in sheets:
            ws = wb['Sheet1']
            h1 = _cell_str(ws, 1, 1)
            h6 = _cell_str(ws, 1, 6)
            h9 = _cell_str(ws, 1, 9)
            h10 = _cell_str(ws, 1, 10)
            h11 = _cell_str(ws, 1, 11)
            h12 = _cell_str(ws, 1, 12)
            h13 = _cell_str(ws, 1, 13)
            h15 = _cell_str(ws, 1, 15)
            h24 = _cell_str(ws, 1, 24)

            # EIBG/EISBG: col 1 = ITEM, col 11 = OTW (flat, Demand only)
            if h1.upper() == 'ITEM' and h11.upper() == 'OTW':
                wb.close()
                return FORMAT_EIBG_EISBG

            # MWC1IPC1: col 1 = PLANT, col 6 = REQUEST ITEM
            if h1.upper() == 'PLANT' and h6.upper() == 'REQUEST ITEM':
                wb.close()
                return FORMAT_MWC1IPC1

            # ICTBG NTL7: col 1 = PLANT, col 10 = REQUEST ITEM (4 rows/part)
            if h1.upper() == 'PLANT' and h10.upper() == 'REQUEST ITEM':
                wb.close()
                return FORMAT_ICTBG_NTL7

            # FMBG: col 1 = PLANT, col 12 = REQUEST ITEM, col 9 = 出貨 (3 rows/part)
            if h1.upper() == 'PLANT' and h12.upper() == 'REQUEST ITEM':
                wb.close()
                return FORMAT_FMBG

            # PSW1+CEW1: col 12 = Status
            if h12 == 'Status':
                wb.close()
                return FORMAT_PSW1_CEW1

            # Weeraya: col 1 = Plant, col 12 = TYPE
            if h1.lower() == 'plant' and h12.upper() == 'TYPE':
                wb.close()
                return FORMAT_WEERAYA

            # ICTBG PSB9 Siriraht: col 15 = REQUEST ITEM (1Demand/2Supply/3Balance)
            if h15.upper() == 'REQUEST ITEM':
                wb.close()
                return FORMAT_ICTBG_PSB9_SIRIRAHT

            # PSBG: col 15 = Filter (values: 1.Demand/2.Supply/3.Net)
            if h15.lower() == 'filter':
                wb.close()
                return FORMAT_PSBG

            # IABG: col 1 = PLANT, col 9 = SHIP, col 13 = PASSDUE (flat)
            if h1.upper() == 'PLANT' and h9.upper() == 'SHIP' and h13.upper() == 'PASSDUE':
                wb.close()
                return FORMAT_IABG

            # Kanyanat: col 24 = TYPE (col 1 = NO, col 3 = Plant)
            if h24.upper() == 'TYPE':
                wb.close()
                return FORMAT_KANYANAT

        wb.close()
        return None
    except Exception:
        return None


def detect_buyer(filepath):
    """
    向後相容: 舊的 detect_buyer() 函式。
    只回傳 'Ketwadee' / 'Kanyanat' / 'Weeraya' / None，其他新格式回傳 None。
    """
    fmt = detect_format(filepath)
    return {
        FORMAT_KETWADEE: 'Ketwadee',
        FORMAT_KANYANAT: 'Kanyanat',
        FORMAT_WEERAYA: 'Weeraya',
    }.get(fmt)


# ---------------------------------------------------------------------------
# 日期欄位工具
# ---------------------------------------------------------------------------

MONTH_NAMES = ('JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
               'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC')

# 每個格式: [(sheet_name, date_start_col), ...]
# date_start_col 是 1-based，從該欄開始掃描至 max_col
# 非日期欄會被 _normalize_date_header 自動過濾掉
_FORMAT_SHEETS = {
    FORMAT_KETWADEE:           [('MRP', 16)],
    FORMAT_KANYANAT:           [('Sheet1', 25)],
    FORMAT_WEERAYA:            [('Sheet1', 14)],
    FORMAT_INDIA_IAI1:         [('PAN JIT', 14)],
    FORMAT_PSW1_CEW1:          [('Sheet1', 14)],
    FORMAT_MWC1IPC1:           [('Sheet1', 9)],
    FORMAT_NBQ1:               [('PAN JIT', 16)],
    FORMAT_SVC1PWC1_DIODE_MOS: [('Diode', 9), ('MOS', 9)],
    FORMAT_PSBG:               [('Sheet1', 16)],
    FORMAT_EIBG_EISBG:         [('Sheet1', 12)],
    FORMAT_FMBG:               [('Sheet1', 16)],
    FORMAT_IABG:               [('Sheet1', 13)],
    FORMAT_ICTBG_NTL7:         [('Sheet1', 13)],
    # FORMAT_ICTBG_PSB9_MRP: sheet 名稱動態 (PSB9_MRP*)，extract_dates 時特殊處理
    FORMAT_ICTBG_PSB9_MRP:     [],
    FORMAT_ICTBG_PSB9_SIRIRAHT:[('Sheet1', 16)],
}


def _get_ictbg_psb9_mrp_sheet(wb):
    """找到 PSB9_MRP* sheet (名稱可能包含日期，例 'PSB9_MRP 0413')"""
    for s in wb.sheetnames:
        if s.startswith('PSB9_MRP'):
            return s
    return wb.sheetnames[0] if wb.sheetnames else None


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
    if not s:
        return None
    s_upper = s.upper()
    if 'PAST' in s_upper or 'PASSDUE' in s_upper or 'OVER DUE' in s_upper:
        return 'PASSDUE'
    if s.isdigit() and len(s) == 8:
        return s
    # MM/DD/YY 或 MM/DD/YYYY (e.g. '03/30/26' → '20260330')
    if '/' in s and ' ' not in s:
        parts = s.split('/')
        if len(parts) == 3:
            try:
                m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
                if 1 <= m <= 12 and 1 <= d <= 31:
                    if y < 100:
                        y += 2000
                    if 2000 <= y <= 2099:
                        return f'{y:04d}{m:02d}{d:02d}'
            except ValueError:
                pass
    # "2026-JUL" → "JUL" (不限定年份，自動適用任何年度)
    if '-' in s and len(s) >= 5:
        parts = s.split('-')
        if len(parts) == 2 and len(parts[1]) == 3:
            month = parts[1].upper()
            if month in MONTH_NAMES:
                return month
    # 直接就是月份名 (JUL, AUG, ...)
    if s_upper in MONTH_NAMES:
        return s_upper
    return None


def _sort_date_cols(dates, anchor_date=None):
    """
    方案二 (匯總格式模式): 固定 26 欄 = PASSDUE + 16 週 + 9 月

    - PASSDUE: 源檔裡既有的必要欄位 (來自 "PASSDUE"/"PAST DUE" 等標籤)，
               直接對應，不會有週日期折入此欄。
    - W1~W16 (K~Z): **以來源檔最早的週日期 (Monday) 為 W1 起點**，產生 16 個
                    連續週一 (YYYYMMDD)。若沒有任何週日期則 fallback 使用 anchor
                    所在週的週一。
    - M1~M9 (AA~AI): 從 W16 次週起算的 9 個月份標籤 (MMM 縮寫)

    來源檔案的日期欄會依據落點映射到固定欄位：
    - 落在 W1~W16 範圍 → 對應週 Monday (若非 Monday 則取該週週一)
    - 落在 M1~M9 範圍 → 對應月份標籤 (多筆會在 reader 累加)
    - 超出 M9 範圍 (太未來) → 丟棄
    - 早於 W1 (理論上不應發生，因為 W1=最早週) → 丟棄並警告
    - 月份標籤不在 M1~M9 → 丟棄

    Args:
        dates: set/iterable of normalized date keys from source files
        anchor_date: datetime, 備援基準 (預設 = datetime.now()，僅在無任何週日期時使用)

    Returns:
        tuple(date_cols, conversions)
        - date_cols: ['PASSDUE', YYYYMMDD×16, MMM×9] 固定 26 欄
        - conversions: dict {原始 key: 轉換後 key 或 None}
    """
    # 1. 找出來源檔最早的週日期 (Monday) 作為 W1 起點
    weekly_source = []
    for d in dates:
        if isinstance(d, str) and d.isdigit() and len(d) == 8:
            try:
                weekly_source.append(datetime.strptime(d, '%Y%m%d'))
            except ValueError:
                pass

    if weekly_source:
        earliest = min(weekly_source)
        # 對齊到該週的週一 (weekday 0 = Monday)
        first_monday = earliest - timedelta(days=earliest.weekday())
        first_monday = first_monday.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        # Fallback: 使用 anchor 所在週的週一
        anchor = anchor_date or datetime.now()
        first_monday = anchor - timedelta(days=anchor.weekday())
        first_monday = first_monday.replace(hour=0, minute=0, second=0, microsecond=0)

    # 2. 固定 16 週 Monday (K~Z)
    weekly_mondays = [first_monday + timedelta(weeks=i) for i in range(16)]
    weekly_keys = [m.strftime('%Y%m%d') for m in weekly_mondays]
    last_monday = weekly_mondays[-1]
    w16_sunday = last_monday + timedelta(days=6)

    # 3. 固定 9 個月份 (AA~AI, 從 W16 次週起算)
    w17_monday = last_monday + timedelta(weeks=1)
    monthly_keys = []
    cy, cm = w17_monday.year, w17_monday.month
    for _ in range(9):
        monthly_keys.append(MONTH_NAMES[cm - 1])
        cm += 1
        if cm > 12:
            cm = 1
            cy += 1
    monthly_key_set = set(monthly_keys)

    final_cols = ['PASSDUE'] + weekly_keys + monthly_keys

    # 4. 建立 conversions
    conversions = {}
    rejected = []
    folded_to_month = {}

    for d in dates:
        if d == 'PASSDUE':
            continue
        if d in MONTH_NAMES:
            if d not in monthly_key_set:
                conversions[d] = None
                rejected.append(d)
            continue
        if isinstance(d, str) and d.isdigit() and len(d) == 8:
            try:
                dt = datetime.strptime(d, '%Y%m%d')
            except ValueError:
                conversions[d] = None
                rejected.append(d)
                continue
            if dt < first_monday:
                # 不應發生 (W1 = 最早週)，但若發生則丟棄並警告
                conversions[d] = None
                rejected.append(d)
            elif dt <= w16_sunday:
                days_from_w1 = (dt - first_monday).days
                week_idx = days_from_w1 // 7
                target_key = weekly_keys[week_idx]
                if target_key != d:
                    conversions[d] = target_key
            else:
                # 超出 W16 → 折疊到對應月份
                m_label = MONTH_NAMES[dt.month - 1]
                if m_label in monthly_key_set:
                    conversions[d] = m_label
                    folded_to_month.setdefault(m_label, []).append(d)
                else:
                    conversions[d] = None
                    rejected.append(d)
        else:
            conversions[d] = None
            rejected.append(d)

    print(f"   📅 方案二起始週 (W1): {weekly_keys[0]}, 結束週 (W16): {weekly_keys[-1]}")
    print(f"   📅 月份範圍 (M1~M9): {monthly_keys[0]} ~ {monthly_keys[-1]}")
    if rejected:
        print(f"   ⚠️ 丟棄超出範圍的日期: {rejected}")
    if folded_to_month:
        for m, src_list in folded_to_month.items():
            print(f"   🔄 折疊到 {m}: {src_list}")

    return final_cols, conversions


def extract_dates_from_files(detected_files):
    """
    從每個檔案的 header 動態提取所有日期欄位 (支援全部 8 種格式)。

    Args:
        detected_files: list of (filepath, format_const) tuples

    Returns:
        tuple(date_cols, per_file_dates, warnings)
        - date_cols: list — 排序後的統一日期欄位 (聯集)
        - per_file_dates: dict — {filename: set(date_keys)}
        - warnings: list — 日期不一致的警告訊息
    """
    per_file_dates = {}

    for filepath, fmt in detected_files:
        sheet_specs = _FORMAT_SHEETS.get(fmt, [])
        file_key = os.path.basename(filepath)

        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        file_dates = set()

        # ICTBG PSB9 MRP 的 sheet 名稱動態 (PSB9_MRP + 日期)
        if fmt == FORMAT_ICTBG_PSB9_MRP:
            sheet_name = _get_ictbg_psb9_mrp_sheet(wb)
            sheet_specs = [(sheet_name, 15)] if sheet_name else []

        for sheet_name, start_col in sheet_specs:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            for col_idx, cell in enumerate(ws[1], start=1):
                if col_idx < start_col:
                    continue
                v = getattr(cell, 'value', None)
                if v is None:
                    continue
                norm = _normalize_date_header(v)
                if norm is not None and norm not in file_dates:
                    file_dates.add(norm)
        wb.close()

        per_file_dates[file_key] = file_dates
        print(f"  {FORMAT_LABELS.get(fmt, fmt)} [{file_key}]: 偵測到 {len(file_dates)} 個日期欄位")

    # 聯集 = 所有檔案的日期
    all_dates = set()
    for dates in per_file_dates.values():
        all_dates |= dates

    # 比對差異 (僅顯示警告，不阻擋)
    warnings = []
    keys = list(per_file_dates.keys())
    for i, k1 in enumerate(keys):
        for k2 in keys[i + 1:]:
            only_in_k1 = per_file_dates[k1] - per_file_dates[k2]
            only_in_k2 = per_file_dates[k2] - per_file_dates[k1]
            if only_in_k1:
                msg = f'{k1} 有但 {k2} 沒有的日期: {sorted(only_in_k1)}'
                warnings.append(msg)
            if only_in_k2:
                msg = f'{k2} 有但 {k1} 沒有的日期: {sorted(only_in_k2)}'
                warnings.append(msg)

    if not warnings:
        print(f"  ✅ 所有檔案日期完全一致 ({len(all_dates)} 個日期)")
    else:
        print(f"  ⚠️ 日期不完全一致，共 {len(warnings)} 個差異，已取聯集 ({len(all_dates)} 個日期)")

    date_cols, conversions = _sort_date_cols(all_dates)
    return date_cols, per_file_dates, warnings, conversions


# 舊名稱保留做向後相容 (app.py 可能還在用)
def extract_dates_from_buyer_files(buyer_files):
    """Deprecated: 使用 extract_dates_from_files()"""
    # 將舊 dict 格式轉換為新 list 格式
    old_format_map = {
        'Ketwadee': FORMAT_KETWADEE,
        'Kanyanat': FORMAT_KANYANAT,
        'Weeraya': FORMAT_WEERAYA,
    }
    detected_files = [
        (fp, old_format_map.get(name, FORMAT_KETWADEE))
        for name, fp in buyer_files.items()
    ]
    return extract_dates_from_files(detected_files)


# ---------------------------------------------------------------------------
# Buyer 讀取器
# ---------------------------------------------------------------------------

def _build_date_col_map(ws, start_col, date_cols, conversions=None):
    """
    建立 column → date_key 的映射。

    方案二支援多個 source 欄位映射到同一個 target key (折疊/累加)：
    - 例: 來源 20260406 + 20260407 都映射到同一個週 Monday → reader 需累加兩者
    - 例: 來源 20261015 + 20261101 都映射到 OCT 月份欄 → reader 需累加

    conversions: dict {原始 key: 轉換後 key 或 None (丟棄)}。
    """
    date_col_map = {}
    date_cols_set = set(date_cols)
    for col_idx, cell in enumerate(ws[1], start=1):
        if col_idx < start_col:
            continue
        v = getattr(cell, 'value', None)
        if v is None:
            continue
        norm = _normalize_date_header(v)
        if norm is None:
            continue
        # 套用轉換 (例: 20261015 → OCT, 20260330 → PASSDUE)
        if conversions and norm in conversions:
            norm = conversions[norm]
            if norm is None:
                continue  # 被丟棄
        if norm in date_cols_set:
            date_col_map[col_idx] = norm
    return date_col_map


def _read_row_dates(row, date_col_map):
    """
    從單一資料 row 讀取日期欄位值，支援累加 (多個 source 欄位 → 同一 target key)。

    Args:
        row: tuple of Cell objects (from openpyxl iter_rows, values_only=False)
        date_col_map: dict {col_idx (1-based): target_date_key}

    Returns:
        dict {target_date_key: accumulated_number}
    """
    data = {}
    for col_idx, date_key in date_col_map.items():
        if col_idx - 1 >= len(row):
            continue
        v = row[col_idx - 1].value
        if v is None or v == '':
            continue
        try:
            v_num = float(v)
        except (ValueError, TypeError):
            continue
        data[date_key] = data.get(date_key, 0) + v_num
    return data


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


def _read_ketwadee(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """讀取 PSB5 Ketwadee: MRP sheet, 3 rows/part"""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['MRP']

    date_col_map = _build_date_col_map(ws, 16, date_cols, conversions)

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

            demand = _read_row_dates(row, date_col_map)
            supply = _read_row_dates(rows[i + 1], date_col_map) if i + 1 < len(rows) else {}

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


def _read_kanyanat(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """讀取 PSB7 Kanyanat: Sheet1, 4 rows/part"""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 25, date_cols, conversions)

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

            demand = _read_row_dates(row, date_col_map)
            supply = _read_row_dates(rows[i + 1], date_col_map) if i + 1 < len(rows) else {}

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


def _read_weeraya(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """讀取 PSB7 Weeraya: Sheet1, 5 rows/part"""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 14, date_cols, conversions)

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

            demand = _read_row_dates(row, date_col_map)
            supply = _read_row_dates(rows[i + 2], date_col_map) if i + 2 < len(rows) else {}
            balance_data = _read_row_dates(rows[i + 3], date_col_map) if i + 3 < len(rows) else {}

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


def _read_india_iai1(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 India IAI1: PAN JIT sheet, 3 rows/part (Demand/Supply/Balance)
    col 3 = PLANT (每列讀), col 4 = PARTNO, col 7 = VENDOR PARTNO,
    col 11 = Stock, col 13 = Request (marker), col 14+ = dates.
    多 PLANT 檔案 — 每列從 col 3 讀 PLANT。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['PAN JIT']

    date_col_map = _build_date_col_map(ws, 14, date_cols, conversions)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        marker = row[12].value if len(row) > 12 else None  # col 13

        if marker == 'Demand':
            row_plant = row[2].value if len(row) > 2 else None  # col 3
            part_no = row[3].value if len(row) > 3 else None     # col 4
            vendor_part = row[6].value if len(row) > 6 else None  # col 7
            stock = row[10].value if len(row) > 10 else 0        # col 11

            demand = _read_row_dates(row, date_col_map)
            supply = _read_row_dates(rows[i + 1], date_col_map) if i + 1 < len(rows) else {}
            balance = _read_row_dates(rows[i + 2], date_col_map) if i + 2 < len(rows) else {}

            results.append({
                'buyer': buyer_label or 'India',
                'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0, 'on_way': None,
                'demand': demand, 'supply': supply,
                'balance_override': balance,
            })
            i += 3
        else:
            i += 1

    wb.close()
    return results


def _read_psw1_cew1(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 PSW1+CEW1: Sheet1, 5 rows/part (A-Demand/B-Supply/C-Net/D-ETD/E-Remark)
    col 3 = PLANT, col 6 = PN, col 8 = MFG (vendor part),
    col 12 = Status (marker), col 13 = STOCK, col 14+ = dates (MM/DD/YY).
    取 A-Demand→Demand, B-Supply→Supply, C-Net→Balance, D/E 跳過。
    多 PLANT 檔案 — 每列從 col 3 讀 PLANT。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 14, date_cols, conversions)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        marker = row[11].value if len(row) > 11 else None  # col 12

        if marker == 'A-Demand':
            row_plant = row[2].value if len(row) > 2 else None   # col 3
            part_no = row[5].value if len(row) > 5 else None      # col 6 = PN
            vendor_part = row[7].value if len(row) > 7 else None  # col 8 = MFG
            stock = row[12].value if len(row) > 12 else 0         # col 13

            demand = _read_row_dates(row, date_col_map)
            supply = _read_row_dates(rows[i + 1], date_col_map) if i + 1 < len(rows) else {}
            balance = _read_row_dates(rows[i + 2], date_col_map) if i + 2 < len(rows) else {}
            # rows[i+3] = D-ETD, rows[i+4] = E-Remark → 跳過

            results.append({
                'buyer': buyer_label or 'PSW1+CEW1',
                'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0, 'on_way': None,
                'demand': demand, 'supply': supply,
                'balance_override': balance,
            })
            i += 5
        else:
            i += 1

    wb.close()
    return results


def _read_mwc1ipc1(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 MWC1+IPC1: Sheet1, 4 rows/part
    (GROSS REQTS/FIRM ORDERS/VENDOR CFM/NET AVAIL)
    col 1 = PLANT, col 2 = PARTNO, col 3 = VENDOR PARTNO,
    col 6 = REQUEST ITEM (marker), col 7 = PLANT STOCK, col 9+ = dates.
    取 GROSS REQTS→Demand, VENDOR CFM→Supply, NET AVAIL→Balance, FIRM ORDERS 跳過。
    多 PLANT 檔案 — 每列從 col 1 讀 PLANT。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 9, date_cols, conversions)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        marker = row[5].value if len(row) > 5 else None  # col 6

        if marker == 'GROSS REQTS':
            row_plant = row[0].value if len(row) > 0 else None    # col 1
            part_no = row[1].value if len(row) > 1 else None       # col 2
            vendor_part = row[2].value if len(row) > 2 else None   # col 3
            stock = row[6].value if len(row) > 6 else 0            # col 7

            demand = _read_row_dates(row, date_col_map)  # GROSS REQTS
            # rows[i+1] = FIRM ORDERS → 跳過
            supply = _read_row_dates(rows[i + 2], date_col_map) if i + 2 < len(rows) else {}   # VENDOR CFM
            balance = _read_row_dates(rows[i + 3], date_col_map) if i + 3 < len(rows) else {}  # NET AVAIL

            results.append({
                'buyer': buyer_label or 'MWC1+IPC1',
                'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0, 'on_way': None,
                'demand': demand, 'supply': supply,
                'balance_override': balance,
            })
            i += 4
        else:
            i += 1

    wb.close()
    return results


def _read_nbq1(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 NBQ1: PAN JIT sheet, 1 row/part
    col 1 = PARTNO, col 3 = VENDOR PARTNO, col 15 = STOCK,
    col 16 = PASSDUE, col 17+ = 週/月日期。
    無 marker column。無 PLANT column → 從檔名比對。
    Demand = 當列日期值, Supply = 空, Balance = 公式 (由產出器處理)。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['PAN JIT']

    date_col_map = _build_date_col_map(ws, 16, date_cols, conversions)

    results = []
    max_col = ws.max_column
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                            min_col=1, max_col=max_col, values_only=False):
        part_no = row[0].value if len(row) > 0 else None       # col 1
        if part_no is None or str(part_no).strip() == '':
            continue
        vendor_part = row[2].value if len(row) > 2 else None   # col 3
        stock = row[14].value if len(row) > 14 else 0          # col 15

        demand = _read_row_dates(row, date_col_map)

        results.append({
            'buyer': buyer_label or 'NBQ1',
            'plant': plant_code or '',
            'part_no': _to_partno(part_no),
            'vendor_part': str(vendor_part) if vendor_part else '',
            'stock': stock or 0, 'on_way': None,
            'demand': demand, 'supply': {},  # Supply 留空
        })

    wb.close()
    return results


def _read_svc1pwc1_diode_mos(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 SVC1+PWC1 DIODE&MOS: 同時處理 Diode 和 MOS 兩個 sheet, 1 row/part
    col 1 = PLANT (每列讀), col 3 = PARTNO, col 5 = VENDOR PARTNO, col 8 = STOCK,
    col 9+ = 日期 (中間夾 NET/SHORTAGE/CFM/出貨/交期 — 非日期欄由 normalizer 自動過濾)。
    Demand = 當列日期值, Supply = 空, Balance = 公式。
    多 PLANT 檔案。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    results = []

    for sheet_name in ('Diode', 'MOS'):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        date_col_map = _build_date_col_map(ws, 9, date_cols, conversions)
        max_col = ws.max_column

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                min_col=1, max_col=max_col, values_only=False):
            row_plant = row[0].value if len(row) > 0 else None    # col 1
            part_no = row[2].value if len(row) > 2 else None       # col 3
            if part_no is None or str(part_no).strip() == '':
                continue
            vendor_part = row[4].value if len(row) > 4 else None   # col 5
            stock = row[7].value if len(row) > 7 else 0            # col 8

            demand = _read_row_dates(row, date_col_map)

            results.append({
                'buyer': buyer_label or 'SVC1+PWC1',
                'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0, 'on_way': None,
                'demand': demand, 'supply': {},  # Supply 留空
            })

    wb.close()
    return results


def _read_psbg(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 PSBG (PSB5 PANJIT): Sheet1, 3 rows/part.
    col 3 = Part No, col 4 = Vendor Part, col 8 = STOCK, col 11 = Ship in Transit,
    col 15 = Filter (1.Demand / 2.Supply / 3.Net), dates from col 16.
    單 PLANT 檔案 — PLANT 從檔名比對。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 16, date_cols, conversions)
    max_col = ws.max_column

    # 每 3 行一組: 1.Demand / 2.Supply / 3.Net
    pending = {}  # part_no → dict
    results = []

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                            min_col=1, max_col=max_col, values_only=False):
        part_no = row[2].value if len(row) > 2 else None       # col 3
        if part_no is None or str(part_no).strip() == '':
            continue
        filter_val = row[14].value if len(row) > 14 else None   # col 15
        fv = str(filter_val).strip() if filter_val else ''

        pn = _to_partno(part_no)

        if fv == '1.Demand':
            vendor_part = row[3].value if len(row) > 3 else None  # col 4
            stock = row[7].value if len(row) > 7 else 0           # col 8
            on_way = row[10].value if len(row) > 10 else None     # col 11
            demand = _read_row_dates(row, date_col_map)
            pending[pn] = {
                'buyer': buyer_label or 'PSBG',
                'plant': plant_code or '',
                'part_no': pn,
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0,
                'on_way': on_way or 0,
                'demand': demand,
                'supply': {},
            }
        elif fv == '2.Supply' and pn in pending:
            pending[pn]['supply'] = _read_row_dates(row, date_col_map)
        elif fv == '3.Net' and pn in pending:
            results.append(pending.pop(pn))

    # 剩餘未完成的 (只有 Demand 沒有 Net)
    results.extend(pending.values())

    wb.close()
    return results


def _read_eibg_eisbg(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 EIBG/EISBG: Sheet1, 1 row/part (flat, Demand only).
    col 1 = ITEM, col 3 = PARTNO, col 4 = VENDOR PARTNO, col 10 = PLANT STOCK,
    col 11 = OTW, col 12+ = dates.
    單 PLANT 檔案 — PLANT 從檔名比對。
    Demand = 當列日期值, Supply = 空, Balance = 公式 (由產出器處理)。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 12, date_cols, conversions)

    results = []
    max_col = ws.max_column
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                            min_col=1, max_col=max_col, values_only=False):
        part_no = row[2].value if len(row) > 2 else None       # col 3
        if part_no is None or str(part_no).strip() == '':
            continue
        vendor_part = row[3].value if len(row) > 3 else None   # col 4
        stock = row[9].value if len(row) > 9 else 0            # col 10 PLANT STOCK
        on_way = row[10].value if len(row) > 10 else None      # col 11 OTW

        demand = _read_row_dates(row, date_col_map)

        results.append({
            'buyer': buyer_label or 'EIBG',
            'plant': plant_code or '',
            'part_no': _to_partno(part_no),
            'vendor_part': str(vendor_part) if vendor_part else '',
            'stock': stock or 0, 'on_way': on_way or 0,
            'demand': demand, 'supply': {},
        })

    wb.close()
    return results


def _read_fmbg(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 FMBG: Sheet1, 3 rows/part (A-Demand/B-CFM/C-Bal).
    col 1 = PLANT, col 5 = PARTNO, col 7 = VENDOR PARTNO, col 12 = REQUEST ITEM,
    col 15 = PLANT STOCK, col 16+ = dates (含 PASSDUE).
    A-Demand → Demand, B-CFM → Supply, C-Bal → Balance。
    多 PLANT 檔案 — 每列從 col 1 讀 PLANT。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 16, date_cols, conversions)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        marker = row[11].value if len(row) > 11 else None  # col 12

        if marker == 'A-Demand':
            row_plant = row[0].value if len(row) > 0 else None    # col 1
            part_no = row[4].value if len(row) > 4 else None       # col 5
            vendor_part = row[6].value if len(row) > 6 else None   # col 7
            stock = row[14].value if len(row) > 14 else 0          # col 15

            demand = _read_row_dates(row, date_col_map)
            supply = _read_row_dates(rows[i + 1], date_col_map) if i + 1 < len(rows) else {}
            balance = _read_row_dates(rows[i + 2], date_col_map) if i + 2 < len(rows) else {}

            results.append({
                'buyer': buyer_label or 'FMBG',
                'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0, 'on_way': None,
                'demand': demand, 'supply': supply,
                'balance_override': balance,
            })
            i += 3
        else:
            i += 1

    wb.close()
    return results


def _read_iabg(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 IABG: Sheet1, 1 row/part (flat, Demand only).
    col 1 = PLANT, col 4 = PARTNO, col 5 = VENDOR PARTNO, col 12 = PLANT STOCK,
    col 13+ = dates (含 PASSDUE).
    多 PLANT 檔案 — 每列從 col 1 讀 PLANT。
    Demand = 當列日期值, Supply = 空, Balance = 公式 (由產出器處理)。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 13, date_cols, conversions)

    results = []
    max_col = ws.max_column
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                            min_col=1, max_col=max_col, values_only=False):
        row_plant = row[0].value if len(row) > 0 else None    # col 1
        part_no = row[3].value if len(row) > 3 else None       # col 4
        if part_no is None or str(part_no).strip() == '':
            continue
        vendor_part = row[4].value if len(row) > 4 else None   # col 5
        stock = row[11].value if len(row) > 11 else 0          # col 12

        demand = _read_row_dates(row, date_col_map)

        results.append({
            'buyer': buyer_label or 'IABG',
            'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
            'part_no': _to_partno(part_no),
            'vendor_part': str(vendor_part) if vendor_part else '',
            'stock': stock or 0, 'on_way': None,
            'demand': demand, 'supply': {},
        })

    wb.close()
    return results


def _read_ictbg_ntl7(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 ICTBG NTL7: Sheet1, 4 rows/part (GROSS REQTS/FIRM ORDERS/Vendor Cfm/NET AVAIL).
    col 1 = PLANT, col 2 = PARTNO, col 7 = VENDOR PARTNO, col 10 = REQUEST ITEM,
    col 11 = PLANT STOCK, col 13+ = dates (含 PASSDUE).
    GROSS REQTS → Demand, Vendor Cfm → Supply, NET AVAIL → Balance, FIRM ORDERS 跳過。
    多 PLANT 檔案 — 每列從 col 1 讀 PLANT。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 13, date_cols, conversions)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        marker = row[9].value if len(row) > 9 else None  # col 10

        if marker == 'GROSS REQTS':
            row_plant = row[0].value if len(row) > 0 else None    # col 1
            part_no = row[1].value if len(row) > 1 else None       # col 2
            vendor_part = row[6].value if len(row) > 6 else None   # col 7
            stock = row[10].value if len(row) > 10 else 0          # col 11

            demand = _read_row_dates(row, date_col_map)
            # rows[i+1] = FIRM ORDERS → 跳過
            supply = _read_row_dates(rows[i + 2], date_col_map) if i + 2 < len(rows) else {}   # Vendor Cfm
            balance = _read_row_dates(rows[i + 3], date_col_map) if i + 3 < len(rows) else {}  # NET AVAIL

            results.append({
                'buyer': buyer_label or 'ICTBG-NTL7',
                'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0, 'on_way': None,
                'demand': demand, 'supply': supply,
                'balance_override': balance,
            })
            i += 4
        else:
            i += 1

    wb.close()
    return results


def _read_ictbg_psb9_mrp(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 ICTBG PSB9 Kaewarin: sheet 名稱 PSB9_MRP*, 4 rows/part (DEMAND/SUPPLY/NET/Remark).
    col 1 = Plant, col 3 = Part No, col 8 = Vendor Part, col 11 = STOCK QTY,
    col 14 = Type, col 15+ = dates (含 PAST DUE).
    DEMAND → Demand, SUPPLY → Supply, NET → Balance, Remark 跳過。
    多 PLANT 檔案 — 每列從 col 1 讀 PLANT。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    sheet_name = _get_ictbg_psb9_mrp_sheet(wb)
    if not sheet_name or sheet_name not in wb.sheetnames:
        wb.close()
        return []
    ws = wb[sheet_name]

    date_col_map = _build_date_col_map(ws, 15, date_cols, conversions)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        marker = row[13].value if len(row) > 13 else None  # col 14 Type
        marker_str = str(marker).strip().upper() if marker else ''

        if marker_str == 'DEMAND':
            row_plant = row[0].value if len(row) > 0 else None    # col 1
            part_no = row[2].value if len(row) > 2 else None       # col 3
            vendor_part = row[7].value if len(row) > 7 else None   # col 8
            stock = row[10].value if len(row) > 10 else 0          # col 11

            demand = _read_row_dates(row, date_col_map)
            # 尋找同料號的 SUPPLY 與 NET 行 (可能有 Remark 在中間或後面)
            supply, balance = {}, {}
            j = i + 1
            next_demand = i + 1
            while j < len(rows) and j < i + 5:
                m = rows[j][13].value if len(rows[j]) > 13 else None
                m_str = str(m).strip().upper() if m else ''
                if m_str == 'DEMAND':
                    next_demand = j
                    break
                if m_str == 'SUPPLY':
                    supply = _read_row_dates(rows[j], date_col_map)
                elif m_str == 'NET':
                    balance = _read_row_dates(rows[j], date_col_map)
                j += 1
                next_demand = j

            results.append({
                'buyer': buyer_label or 'ICTBG-PSB9',
                'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0, 'on_way': None,
                'demand': demand, 'supply': supply,
                'balance_override': balance,
            })
            i = next_demand
        else:
            i += 1

    wb.close()
    return results


def _read_ictbg_psb9_siriraht(filepath, date_cols, buyer_label=None, plant_code=None, conversions=None):
    """
    讀取 ICTBG PSB9 Siriraht: Sheet1, 3 rows/part (1Demand/2Supply/3Balance).
    col 1 = PLANT, col 4 = PARTNO, col 6 = VENDOR PARTNO, col 12 = PSB9 STOCK,
    col 15 = REQUEST ITEM, col 16+ = dates (含 PASSDUE).
    多 PLANT 檔案 — 每列從 col 1 讀 PLANT。
    """
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb['Sheet1']

    date_col_map = _build_date_col_map(ws, 16, date_cols, conversions)

    results = []
    max_col = ws.max_column
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row,
                             min_col=1, max_col=max_col, values_only=False))
    i = 0
    while i < len(rows):
        row = rows[i]
        marker = row[14].value if len(row) > 14 else None  # col 15
        marker_str = str(marker).strip() if marker else ''

        if marker_str == '1Demand':
            row_plant = row[0].value if len(row) > 0 else None    # col 1
            part_no = row[3].value if len(row) > 3 else None       # col 4
            vendor_part = row[5].value if len(row) > 5 else None   # col 6
            stock = row[11].value if len(row) > 11 else 0          # col 12

            demand = _read_row_dates(row, date_col_map)
            supply = _read_row_dates(rows[i + 1], date_col_map) if i + 1 < len(rows) else {}
            balance = _read_row_dates(rows[i + 2], date_col_map) if i + 2 < len(rows) else {}

            results.append({
                'buyer': buyer_label or 'ICTBG-PSB9-S',
                'plant': str(row_plant).strip() if row_plant else (plant_code or ''),
                'part_no': _to_partno(part_no),
                'vendor_part': str(vendor_part) if vendor_part else '',
                'stock': stock or 0, 'on_way': None,
                'demand': demand, 'supply': supply,
                'balance_override': balance,
            })
            i += 3
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

FORMAT_READERS = {
    FORMAT_KETWADEE:           _read_ketwadee,
    FORMAT_KANYANAT:           _read_kanyanat,
    FORMAT_WEERAYA:            _read_weeraya,
    FORMAT_INDIA_IAI1:         _read_india_iai1,
    FORMAT_PSW1_CEW1:          _read_psw1_cew1,
    FORMAT_MWC1IPC1:           _read_mwc1ipc1,
    FORMAT_NBQ1:               _read_nbq1,
    FORMAT_SVC1PWC1_DIODE_MOS: _read_svc1pwc1_diode_mos,
    FORMAT_PSBG:               _read_psbg,
    FORMAT_EIBG_EISBG:         _read_eibg_eisbg,
    FORMAT_FMBG:               _read_fmbg,
    FORMAT_IABG:               _read_iabg,
    FORMAT_ICTBG_NTL7:         _read_ictbg_ntl7,
    FORMAT_ICTBG_PSB9_MRP:     _read_ictbg_psb9_mrp,
    FORMAT_ICTBG_PSB9_SIRIRAHT:_read_ictbg_psb9_siriraht,
}

# 向後相容: 舊 API
BUYER_READERS = {
    'Ketwadee': _read_ketwadee,
    'Kanyanat': _read_kanyanat,
    'Weeraya':  _read_weeraya,
}


def detect_all_formats(forecast_files):
    """
    Pre-detection: 偵測所有檔案的格式，一次回傳所有識別結果。
    Returns:
        tuple(detected, unknown)
        - detected: list of (filepath, format_const) tuples (成功識別)
        - unknown: list of filepaths (無法識別)
    """
    detected, unknown = [], []
    for fp in forecast_files:
        fmt = detect_format(fp)
        if fmt is None:
            unknown.append(fp)
        else:
            detected.append((fp, fmt))
    return detected, unknown


def consolidate(forecast_files, reference_template, output_path,
                erp_mapping=None, plant_codes=None, file_labels=None):
    """
    合併多個 Delta Forecast 檔案為匯總格式 Excel (支援 8 種格式)。

    Args:
        forecast_files: list of file paths (1 個或多個，任何格式組合)
        reference_template: 匯總格式模板路徑 (用於取得表頭格式，不用於日期)
        output_path: 輸出檔案路徑
        erp_mapping: dict {Plant: (ERP客戶簡稱, ERP送貨地點)}, 已棄用
        plant_codes: list of valid PLANT codes, 用於從檔名比對單 PLANT 檔案的 PLANT。
                     建議由 customer_mappings 的 region 欄位提取
                     (例: ['PSB5', 'PSB7', 'IAI1', 'IPC1'])。
        file_labels: dict {filepath: label}，用於自訂 Buyer 欄位顯示名稱。
                     通常由 app.py 傳入原始檔名 (因為暫存檔名被改成 forecast_temp_N)。
                     若未提供則使用 basename without ext。

    Returns:
        dict with keys: success, part_count, format_stats, unknown_files,
                        date_warnings, message
    """
    if not forecast_files:
        return {
            'success': False, 'part_count': 0,
            'message': '未提供任何 Forecast 檔案'
        }

    # 1. 預先偵測所有檔案格式 (任一失敗即回報所有失敗檔案)
    detected, unknown = detect_all_formats(forecast_files)
    if unknown:
        unknown_names = [os.path.basename(fp) for fp in unknown]
        return {
            'success': False, 'part_count': 0,
            'unknown_files': unknown_names,
            'message': '無法識別以下檔案格式 (請確認為 Delta 8 種標準格式): '
                       + ', '.join(unknown_names)
        }

    print(f"Delta 合併: 偵測到 {len(detected)} 個檔案")
    for fp, fmt in detected:
        print(f"  [{FORMAT_LABELS.get(fmt, fmt)}] {os.path.basename(fp)}")

    # 2. 從所有檔案動態提取日期欄位 (取聯集)
    #    conversions: {原始 YYYYMMDD: 月份標籤 或 None}，讀取器要用這個 map
    #    把來源檔的月末日期轉換為月份標籤 (避免週/月欄位重複)。
    date_cols, per_file_dates, date_warnings, conversions = extract_dates_from_files(detected)
    if not date_cols:
        return {
            'success': False, 'part_count': 0,
            'message': '無法從 Forecast 檔案提取日期欄位'
        }

    print(f"  統一日期欄位: {len(date_cols)} 個 ({date_cols[0]} ~ {date_cols[-1]})")
    if conversions:
        non_null_conv = {k: v for k, v in conversions.items() if v is not None}
        rejected_conv = [k for k, v in conversions.items() if v is None]
        if non_null_conv:
            print(f"  月末→月份轉換: {non_null_conv}")
        if rejected_conv:
            print(f"  拒絕日期: {rejected_conv}")

    # 3. 讀取各檔案資料
    # Buyer 欄位暫時使用檔案名稱 (不含副檔名)，等客戶確認代碼→名稱對照後再調整邏輯。
    # PLANT 欄位:
    #   - 單 PLANT 檔案: 從檔名比對 plant_codes (例: 'PSB5 Ketwadee.xlsx' → 'PSB5')
    #   - 多 PLANT 檔案: 每列從工作表欄位讀取 (不使用檔名)
    all_source = []
    format_stats = {}  # {filename: count}

    for fp, fmt in detected:
        reader = FORMAT_READERS.get(fmt)
        if reader is None:
            print(f"  ⚠️ 跳過未註冊 reader 的格式: {fmt} ({os.path.basename(fp)})")
            continue

        # Buyer 顯示名稱 & 檔名比對用途:
        # 優先使用呼叫端傳入的原始檔名 (因為上傳時 app.py 會把檔案改名成 forecast_temp_N)
        label_for_match = file_labels.get(fp) if file_labels else None
        if label_for_match:
            buyer_label = os.path.splitext(label_for_match)[0]
        else:
            buyer_label = os.path.splitext(os.path.basename(fp))[0]
        file_key = os.path.basename(fp)

        # 單 PLANT 檔案才需從檔名比對 PLANT 代碼 (用原始檔名，不用 temp 檔名)
        plant_code = None
        if fmt in SINGLE_PLANT_FORMATS and plant_codes:
            match_target = label_for_match if label_for_match else fp
            matched = match_plants_in_filename(match_target, plant_codes)
            if matched:
                plant_code = matched[0]
                if len(matched) > 1:
                    print(f"  ⚠️ {file_key}: 檔名中有多個 PLANT {matched}, 取 {plant_code}")

        try:
            data = reader(fp, date_cols,
                          buyer_label=buyer_label, plant_code=plant_code,
                          conversions=conversions)
        except Exception as e:
            return {
                'success': False, 'part_count': 0,
                'message': f'讀取檔案失敗 [{file_key}] ({FORMAT_LABELS.get(fmt, fmt)}): {e}'
            }

        # PLANT 統一驗證: 不在 mapping 表 (plant_codes) 內的一律設為空白
        # 適用所有 reader (單 PLANT / 多 PLANT) — 確保 PLANT 欄完全來自 mapping 表
        if plant_codes:
            valid_upper = {str(p).strip().upper() for p in plant_codes if p}
            invalid_seen = set()
            for item in data:
                p = item.get('plant') or ''
                if p and str(p).strip().upper() not in valid_upper:
                    invalid_seen.add(str(p).strip())
                    item['plant'] = ''
            if invalid_seen:
                print(f"  ⚠️ {file_key}: PLANT 不在 mapping 表，已留空: {sorted(invalid_seen)}")

        format_stats[file_key] = len(data)
        all_source.extend(data)

        if fmt in SINGLE_PLANT_FORMATS:
            # 單 PLANT 顯示: 從 data 看實際使用的 PLANT (可能因 mapping 驗證被清空)
            actual = next((d.get('plant') for d in data if d.get('plant')), None)
            plant_display = actual or '(未比對到)'
            print(f"  {file_key} [{FORMAT_LABELS.get(fmt, fmt)}]: "
                  f"{len(data)} 個料號, PLANT={plant_display}")
        else:
            unique_plants = sorted({d.get('plant', '') for d in data if d.get('plant')})
            print(f"  {file_key} [{FORMAT_LABELS.get(fmt, fmt)}]: "
                  f"{len(data)} 個料號, 多 PLANT={unique_plants}")

    if not all_source:
        return {
            'success': False, 'part_count': 0,
            'message': '未讀取到任何料號資料 (檔案可能為空或格式異常)'
        }

    # 4. 產出匯總格式 Excel（日期表頭由 date_cols 動態產生）
    part_count = _generate_consolidated_excel(
        all_source, output_path, reference_template, date_cols, erp_mapping
    )

    print(f"  合併完成: {part_count} 個料號 → {os.path.basename(output_path)}")

    result = {
        'success': True,
        'part_count': part_count,
        'date_col_count': len(date_cols),
        'format_stats': format_stats,
        # 向後相容: buyer_stats 用舊的 3 Buyer 名稱 filter
        'buyer_stats': {
            name: sum(cnt for fk, cnt in format_stats.items() if name.lower() in fk.lower())
            for name in ('Ketwadee', 'Kanyanat', 'Weeraya')
        },
        'message': f'成功合併 {part_count} 個料號 ({len(detected)} 個檔案)'
    }
    if date_warnings:
        result['date_warnings'] = date_warnings
    return result
