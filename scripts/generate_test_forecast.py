# -*- coding: utf-8 -*-
"""
scripts/generate_test_forecast.py

從桌面的台達業務買方檔案中，提取真實的 (plant, partno) 資料，
產生 tests/fixtures/forecast_result_test.xlsx，供測試使用。

Supply 值來源：買方檔案本身的 Supply 列（若有）；否則用 Demand 的 50% 模擬。

用法：
    python scripts/generate_test_forecast.py
"""
import sys
import os
from pathlib import Path
from collections import defaultdict

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

import openpyxl
from delta_unified_reader import read_buyer_file

DESKTOP_DIR = Path(r"C:/Users/petty/Desktop/客戶相關資料/01.強茂/台達業務")
OUTPUT_PATH = ROOT / "tests" / "fixtures" / "forecast_result_test.xlsx"

# 所有需要讀取的買方檔案（fname, plant_codes）
# 跳過：已回填版本、邏輯/彙總目錄
BUYER_FILES = [
    ("EIBG-TPW1\u2014Lydia--0427.xlsx",                           ["TPW1"]),
    ("EIBG-UPW1 PANJIT 0413.xlsx",                                ["UPW1"]),
    ("EISBG-.xlsx",                                               ["UPW1"]),
    ("ICTBG(DNI)-NTL7  4.13 MRP CFM.xlsx",                       ["NTL7"]),
    ("ICTBG-PSB9-Kaewarin_20260413.xlsx",                         ["PSB9"]),
    ("ICTBG-PSB9-Siriraht_20260411.xlsx",                         ["PSB9"]),
    ("FMBG-MRP(TPC5)-100109-2026-4-15.xlsx",                      ["TPC5"]),
    ("IABG-IMW1-\u9648\u59ff\u5bb9_20260413.xlsx",                ["IMW1"]),
    ("MRP(SVC1PWC1 DIODE&MOS)-100109-2026-3-30.xlsx",             ["SVC1", "PWC1"]),
    ("NBQ1.xlsx",                                                 ["NBQ1"]),
    ("PSBG (India IAI1&UPI2&DFI1 DIODES)-Jack0401.xlsx",          ["IAI1", "UPI2", "DFI1"]),
    ("PSBG PSW1+CEW1- \u694a\u6d0b\u5f59\u6574 0330(\u5b8c\u6210) (002).xlsx", ["PSW1", "CEW1"]),
    ("PSBG PSW1+CEW1\u5408\u4f75-Aviva_20260416.xlsx",            ["PSW1", "CEW1"]),
    ("PSBG-PSB7PAN JIT YTMDS APR 20 2026 Kanyanat.xlsx",          ["PSB7"]),
    ("PSBG\xa0PSB7_Kanyanat.S0406(\u5b8c\u6210).xlsx",            ["PSB7"]),
    ("PSBG\xa0PSB5-\xa0Ketwadee0406(\u5b8c\u6210).xlsx",          ["PSB5"]),
    ("PSBG\xa0PSB7-Weeraya0406(\u5b8c\u6210).xlsx",               ["PSB7"]),
    ("W4-PSBG DNI-MWC1&IPC1 MRP 04.20.2026\u738b\u8ff0\u9023.xlsx",            ["MWC1", "IPC1"]),
    ("W4-PSBG DNI-MWC1+IPC1 \u5f37\u8302 MRP+SHIP 2026-4-17\u5468\u6843\u6625.xlsx", ["MWC1", "IPC1"]),
    ("W4-PSBG DNI-MWC1-IPC1-MWT-IPT-100109-\u5f37\u8302-0420MRP \u7f85\u5a1f.xlsx", ["MWC1", "IPC1"]),
    ("\u5f37\u8302 MWC1IPC1 MRP 03.30.2026.xlsx",                 ["MWC1", "IPC1"]),
]

MONTH_ORDER = ['PASSDUE', 'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
               'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']


def date_key_sort(key):
    """排序：PASSDUE → 8位數日期 → 月份縮寫"""
    if key == 'PASSDUE':
        return (0, key)
    if len(key) == 8 and key.isdigit():
        return (1, key)
    try:
        return (2, str(MONTH_ORDER.index(key)).zfill(2))
    except ValueError:
        return (3, key)


def main():
    print(f"讀取買方檔案來源: {DESKTOP_DIR}")
    print(f"輸出路徑: {OUTPUT_PATH}")
    print()

    # (plant_upper, partno) → {date_key: supply_val}
    supply_data = defaultdict(lambda: defaultdict(float))
    # (plant_upper, partno) → {date_key: demand_val}
    demand_data = defaultdict(lambda: defaultdict(float))
    # (plant_upper, partno) → (buyer_name, vendor_part, stock, on_way)
    meta = {}

    all_date_keys = set()
    n_files_ok = 0
    n_files_skip = 0
    n_rows_total = 0

    for fname, plant_codes in BUYER_FILES:
        fpath = DESKTOP_DIR / fname
        if not fpath.exists():
            short = (fname[:40] if len(fname) > 40 else fname).encode('ascii','replace').decode()
            print(f"  [SKIP] not found: {short}")
            n_files_skip += 1
            continue

        try:
            rows = read_buyer_file(str(fpath), plant_codes=plant_codes or None,
                                   file_label=fname)
        except Exception as e:
            short = (fname[:40] if len(fname) > 40 else fname).encode('ascii','replace').decode()
            print(f"  [ERR]  {short}: {e}")
            n_files_skip += 1
            continue

        if not rows:
            short = (fname[:40] if len(fname) > 40 else fname).encode('ascii','replace').decode()
            print(f"  [SKIP] 0 rows: {short}")
            n_files_skip += 1
            continue

        n_files_ok += 1
        n_rows_total += len(rows)
        short = fname[:40] if len(fname) > 40 else fname
        print(f"  [OK]   {short.encode('ascii','replace').decode()}: {len(rows)} rows")

        for r in rows:
            partno = str(r.get('part_no', '')).strip()
            plant  = str(r.get('plant', '')).strip().upper()
            if not partno or not plant:
                continue

            key = (plant, partno)

            # 收集 meta（只記錄第一次）
            if key not in meta:
                meta[key] = (
                    fname,
                    str(r.get('vendor_part', '') or ''),
                    r.get('stock') or 0,
                    r.get('on_way') or 0,
                )

            # 收集 demand
            demand = r.get('demand') or {}
            for dk, dv in demand.items():
                if dv:
                    demand_data[key][dk] += float(dv)
                    all_date_keys.add(dk)

            # 收集 supply（優先用實際 supply；否則後面用 demand 的 50%）
            supply = r.get('supply') or {}
            for sk, sv in supply.items():
                if sv:
                    supply_data[key][sk] += float(sv)
                    all_date_keys.add(sk)

    print()
    print(f"讀取完成: {n_files_ok} 個檔案成功, {n_files_skip} 個跳過")
    print(f"唯一 (plant, partno) 組合: {len(meta)}")
    print(f"唯一日期 key: {len(all_date_keys)}")

    # 排序日期 key
    sorted_dates = sorted(all_date_keys, key=date_key_sort)

    # 對 supply 為空的 partno，用 demand 50% 填入（確保測試有值可寫）
    n_demand_fallback = 0
    for key in meta:
        if not supply_data[key]:
            d = demand_data[key]
            if d:
                for dk, dv in d.items():
                    if dv and dv > 0:
                        supply_data[key][dk] = round(dv * 0.5)
                n_demand_fallback += 1

    print(f"Supply 來自實際買方資料: {len(supply_data) - n_demand_fallback} 筆")
    print(f"Supply 以 Demand×50% 推算: {n_demand_fallback} 筆")

    # ── 建立 forecast_result.xlsx ──────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Header（對齊 delta_original_backfill.py 的常數定義）
    # FR_COL_PLANT=2, FR_COL_PARTNO=5, FR_COL_ROW_TYPE=9, FR_DATE_START_COL=10
    header = [
        'Buyer',           # A col 1
        'PLANT',           # B col 2
        'ERP客戶',         # C col 3
        'ERP廠址',         # D col 4
        'PARTNO',          # E col 5
        'VENDOR PARTNO',   # F col 6
        'STOCK',           # G col 7
        'ON-WAY',          # H col 8
        'Date',            # I col 9  (row type: Demand/Supply)
        'PASSDUE',         # J col 10 - first date (hardcode PASSDUE here as non-date header)
    ]
    # 加入其他日期欄（除了 PASSDUE，因為已加在 header[9]）
    other_dates = [d for d in sorted_dates if d != 'PASSDUE']
    header += other_dates
    ws.append(header)

    # 寫資料列
    n_demand_rows = 0
    n_supply_rows = 0
    for key in sorted(meta.keys()):
        plant, partno = key
        buyer_name, vendor_part, stock, on_way = meta[key]

        demand_dict = demand_data[key]
        supply_dict = supply_data[key]

        def get_val(d, dk):
            v = d.get(dk)
            if v is None:
                return None
            return int(v) if isinstance(v, float) and v.is_integer() else v

        # Demand row
        demand_row = [buyer_name, plant, plant, plant + '_SH', partno,
                      vendor_part, stock, on_way, 'Demand']
        demand_row.append(get_val(demand_dict, 'PASSDUE'))
        for dk in other_dates:
            demand_row.append(get_val(demand_dict, dk))
        ws.append(demand_row)
        n_demand_rows += 1

        # Supply row
        supply_row = [buyer_name, plant, plant, plant + '_SH', partno,
                      vendor_part, None, None, 'Supply']
        supply_row.append(get_val(supply_dict, 'PASSDUE'))
        for dk in other_dates:
            supply_row.append(get_val(supply_dict, dk))
        ws.append(supply_row)
        n_supply_rows += 1

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(OUTPUT_PATH))

    print()
    print(f"寫入完成: {n_demand_rows} Demand 列 + {n_supply_rows} Supply 列")
    print(f"輸出: {OUTPUT_PATH}")
    print()
    print("[OK] Done. Update FORECAST_RESULT in tests to point to this file.")


if __name__ == '__main__':
    main()
