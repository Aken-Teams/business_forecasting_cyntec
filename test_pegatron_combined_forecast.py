# -*- coding: utf-8 -*-
"""
Test Pegatron Transit + ERP -> Forecast 整合測試
將在途的 ETA QTY 和 ERP 的淨需求一起填入 Forecast 對應日期欄位
使用 pywin32 COM 保留格式和公式
"""
import pandas as pd
import sys
import os
import shutil
import pythoncom
from win32com import client as win32
from datetime import datetime, timedelta
sys.stdout.reconfigure(encoding='utf-8')

# ==================== ERP 日期計算函數 ====================

def get_week_end_by_breakpoint(schedule_date, breakpoint_text):
    """
    根據排程出貨日期和斷點計算該斷點週的結束日
    斷點就是週的結束日
    """
    breakpoint_map = {
        '週一': 0, '週二': 1, '週三': 2, '週四': 3,
        '週五': 4, '週六': 5, '週日': 6,
        '禮拜一': 0, '禮拜二': 1, '禮拜三': 2, '禮拜四': 3,
        '禮拜五': 4, '禮拜六': 5, '禮拜日': 6,
    }

    breakpoint_weekday = breakpoint_map.get(breakpoint_text, 2)
    current_weekday = schedule_date.weekday()

    if current_weekday <= breakpoint_weekday:
        days_to_breakpoint = breakpoint_weekday - current_weekday
    else:
        days_to_breakpoint = 7 - (current_weekday - breakpoint_weekday)

    week_end = schedule_date + timedelta(days=days_to_breakpoint)
    return week_end

def calculate_erp_eta_target_date(week_end, eta_text):
    """
    根據 ETA 文字和斷點週結束日計算目標日期
    ETA 格式: 本週X, 下週X, 下下週X
    """
    eta_weekday_map = {
        '一': 0, '二': 1, '三': 2, '四': 3,
        '五': 4, '六': 5, '日': 6, '天': 6,
    }

    if not eta_text or pd.isna(eta_text):
        return None

    eta_text = str(eta_text).strip()

    if '下下週' in eta_text or '下下周' in eta_text:
        weeks_offset = 2
        weekday_char = eta_text.replace('下下週', '').replace('下下周', '').strip()
    elif '下週' in eta_text or '下周' in eta_text:
        weeks_offset = 1
        weekday_char = eta_text.replace('下週', '').replace('下周', '').strip()
    elif '本週' in eta_text or '本周' in eta_text:
        weeks_offset = 0
        weekday_char = eta_text.replace('本週', '').replace('本周', '').strip()
    else:
        return None

    target_weekday = eta_weekday_map.get(weekday_char, 1)
    week_end_weekday = week_end.weekday()

    days_diff = target_weekday - week_end_weekday
    target_date = week_end + timedelta(days=7 * weeks_offset + days_diff)

    return target_date

# ==================== 共用函數 ====================

def find_week_column(target_date, date_columns):
    """
    找到目標日期所在週的欄位
    Forecast 日期已經是週一
    """
    if isinstance(target_date, datetime):
        target_date = target_date.date()

    days_since_monday = target_date.weekday()
    week_monday = target_date - timedelta(days=days_since_monday)

    if week_monday in date_columns:
        return date_columns[week_monday]

    for week_start, col_idx in date_columns.items():
        week_end = week_start + timedelta(days=6)
        if week_start <= target_date <= week_end:
            return col_idx

    return None

def test_combined_forecast():
    print("=" * 70)
    print("Test Pegatron Transit + ERP -> Forecast 整合測試")
    print("=" * 70)

    # 檔案路徑
    input_forecast = r'd:\github\business_forecasting_pc\test\cleaned_forecast.xls'
    output_forecast = r'd:\github\business_forecasting_pc\test\forecast_combined.xls'
    transit_file = r'd:\github\business_forecasting_pc\test\在途.xlsx'
    erp_file = r'd:\github\business_forecasting_pc\test\integrated_erp.xlsx'

    # 1. 讀取資料
    forecast_df = pd.read_excel(input_forecast, header=None)
    print(f"Forecast 行數: {len(forecast_df)}, 欄數: {len(forecast_df.columns)}")

    transit_df = pd.read_excel(transit_file)
    print(f"Transit 行數: {len(transit_df)}")

    erp_df = pd.read_excel(erp_file)
    print(f"ERP 行數: {len(erp_df)}")

    # 2. 建立 Forecast 區塊結構
    forecast_blocks = []
    row_idx = 2
    while row_idx < len(forecast_df):
        m_val = forecast_df.iloc[row_idx, 12] if pd.notna(forecast_df.iloc[row_idx, 12]) else ''
        if m_val == 'WEEK#':
            f_val = str(forecast_df.iloc[row_idx, 5]).strip() if pd.notna(forecast_df.iloc[row_idx, 5]) else ''
            g_val = str(forecast_df.iloc[row_idx, 6]).strip() if pd.notna(forecast_df.iloc[row_idx, 6]) else ''
            line_po = f"{f_val}-{g_val}" if f_val and g_val else ''

            ordered_item = ''
            if row_idx + 1 < len(forecast_df):
                ordered_item = str(forecast_df.iloc[row_idx + 1, 8]).strip() if pd.notna(forecast_df.iloc[row_idx + 1, 8]) else ''

            eta_qty_row = row_idx + 4  # ETA QTY 行 (Transit 和 ERP 都填這裡)

            forecast_blocks.append({
                'start_row': row_idx,
                'line_po': line_po,
                'ordered_item': ordered_item,
                'eta_qty_row': eta_qty_row + 1,  # Excel 1-based
            })
            row_idx += 8
        else:
            row_idx += 1

    print(f"\n找到 {len(forecast_blocks)} 個 Forecast 區塊:")
    for block in forecast_blocks:
        print(f"  Line PO={block['line_po']}, Ordered Item={block['ordered_item']}, ETA QTY Row={block['eta_qty_row']}")

    # 3. 取得日期欄位對應
    date_columns = {}
    for col_idx in range(14, len(forecast_df.columns)):
        date_val = forecast_df.iloc[1, col_idx]
        if pd.notna(date_val):
            if isinstance(date_val, str):
                try:
                    date_obj = pd.to_datetime(date_val)
                    date_columns[date_obj.date()] = col_idx + 1
                except:
                    pass
            elif isinstance(date_val, (datetime, pd.Timestamp)):
                date_columns[date_val.date()] = col_idx + 1

    print(f"\n日期欄位對應: {len(date_columns)} 個日期")

    # ==================== 處理 Transit ====================
    print("\n" + "=" * 50)
    print("處理 Transit 資料 (填入 ETA QTY 行)")
    print("=" * 50)

    transit_updates = []
    transit_with_line = transit_df[transit_df['Line 客戶採購單號'].notna()]
    print(f"有 Line 客戶採購單號 的在途記錄: {len(transit_with_line)}")

    for idx, transit_row in transit_with_line.iterrows():
        transit_line_po = str(transit_row['Line 客戶採購單號']).strip()
        transit_ordered_item = str(transit_row['Ordered Item']).strip()
        transit_qty = transit_row['Qty'] if pd.notna(transit_row['Qty']) else 0
        transit_eta = transit_row['ETA']

        matched_block = None
        for block in forecast_blocks:
            if block['line_po'] == transit_line_po and block['ordered_item'] == transit_ordered_item:
                matched_block = block
                break

        if matched_block and pd.notna(transit_eta):
            eta_date = pd.to_datetime(transit_eta).date()
            days_since_monday = eta_date.weekday()
            week_start = eta_date - timedelta(days=days_since_monday)

            if week_start in date_columns:
                excel_col = date_columns[week_start]
                eta_qty_value = transit_qty * 1000

                print(f"\nTransit: Line PO={transit_line_po}, Ordered Item={transit_ordered_item}")
                print(f"  Qty={transit_qty}, ETA={eta_date}")
                print(f"  -> ETA QTY Row {matched_block['eta_qty_row']}, Col {excel_col}, 值 = {eta_qty_value}")

                transit_updates.append((matched_block['eta_qty_row'], excel_col, eta_qty_value))

    print(f"\nTransit 更新筆數: {len(transit_updates)}")

    # ==================== 處理 ERP ====================
    print("\n" + "=" * 50)
    print("處理 ERP 資料 (填入 FORECAST 行)")
    print("=" * 50)

    erp_updates = []
    erp_with_mapping = erp_df[erp_df['客戶需求地區'].notna() & (erp_df['客戶需求地區'] != '')]
    print(f"有 mapping 的 ERP 記錄: {len(erp_with_mapping)}")

    for idx, erp_row in erp_with_mapping.iterrows():
        erp_line_po = str(erp_row['Line 客戶採購單號']).strip() if pd.notna(erp_row['Line 客戶採購單號']) else ''
        erp_pn = str(erp_row['客戶料號']).strip() if pd.notna(erp_row['客戶料號']) else ''
        erp_qty = erp_row['淨需求'] if pd.notna(erp_row['淨需求']) else 0
        erp_schedule_date = erp_row['排程出貨日期']
        erp_breakpoint = str(erp_row['排程出貨日期斷點']).strip() if pd.notna(erp_row['排程出貨日期斷點']) else ''
        erp_eta = str(erp_row['ETA']).strip() if pd.notna(erp_row['ETA']) else ''

        if not erp_line_po or not erp_pn or erp_qty == 0:
            continue

        matched_block = None
        for block in forecast_blocks:
            if block['line_po'] == erp_line_po and block['ordered_item'] == erp_pn:
                matched_block = block
                break

        if not matched_block:
            continue

        if pd.isna(erp_schedule_date) or not erp_breakpoint or not erp_eta:
            continue

        schedule_date = pd.to_datetime(erp_schedule_date)
        week_end = get_week_end_by_breakpoint(schedule_date, erp_breakpoint)
        target_date = calculate_erp_eta_target_date(week_end, erp_eta)

        if target_date is None:
            continue

        excel_col = find_week_column(target_date, date_columns)

        if excel_col is None:
            continue

        forecast_value = erp_qty * 1000

        print(f"\nERP: Line PO={erp_line_po}, PN={erp_pn}, Qty={erp_qty}")
        print(f"  排程={schedule_date.date()}, 斷點={erp_breakpoint}, ETA={erp_eta}")
        print(f"  -> 目標日期={target_date.date()} -> ETA QTY Row {matched_block['eta_qty_row']}, Col {excel_col}, 值 = {forecast_value}")

        erp_updates.append((matched_block['eta_qty_row'], excel_col, forecast_value))

    print(f"\nERP 更新筆數: {len(erp_updates)}")

    # ==================== 合併更新並寫入 ====================
    all_updates = transit_updates + erp_updates

    if not all_updates:
        print("\n沒有需要更新的資料")
        return

    print(f"\n=== 使用 COM 更新 {len(all_updates)} 個儲存格 ===")

    # 複製檔案
    shutil.copy2(input_forecast, output_forecast)
    abs_path = os.path.abspath(output_forecast)

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(abs_path)
        ws = wb.Sheets(1)

        # 合併相同位置的值（累加）
        update_dict = {}
        for row, col, value in all_updates:
            key = (row, col)
            if key in update_dict:
                update_dict[key] += value
            else:
                update_dict[key] = value

        # 更新儲存格
        for (row, col), value in update_dict.items():
            current_val = ws.Cells(row, col).Value
            if current_val is None or current_val == '' or current_val == 0:
                ws.Cells(row, col).Value = value
            else:
                ws.Cells(row, col).Value = float(current_val) + value
            print(f"  更新 Row {row}, Col {col} = {ws.Cells(row, col).Value}")

        # 保存
        wb.SaveCopyAs(abs_path + ".tmp")
        wb.Close(SaveChanges=False)
        wb = None

        if os.path.exists(abs_path):
            os.remove(abs_path)
        os.rename(abs_path + ".tmp", abs_path)

        print(f"\n已輸出到: {output_forecast}")

        # 統計
        print("\n" + "=" * 50)
        print("統計")
        print("=" * 50)
        print(f"Transit 更新: {len(transit_updates)} 筆")
        print(f"ERP 更新: {len(erp_updates)} 筆")
        print(f"合併後位置: {len(update_dict)} 個")

    except Exception as e:
        print(f"錯誤: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    test_combined_forecast()
