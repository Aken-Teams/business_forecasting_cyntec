# -*- coding: utf-8 -*-
"""
Test Pegatron ERP -> Forecast 邏輯
將 ERP 的淨需求填入 Forecast 對應日期欄位
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

def get_week_end_by_breakpoint(schedule_date, breakpoint_text):
    """
    根據排程出貨日期和斷點計算該斷點週的結束日
    斷點就是週的結束日
    """
    # 斷點對應的星期幾 (0=週一, 6=週日)
    breakpoint_map = {
        '週一': 0, '週二': 1, '週三': 2, '週四': 3,
        '週五': 4, '週六': 5, '週日': 6,
        '禮拜一': 0, '禮拜二': 1, '禮拜三': 2, '禮拜四': 3,
        '禮拜五': 4, '禮拜六': 5, '禮拜日': 6,
    }

    breakpoint_weekday = breakpoint_map.get(breakpoint_text, 2)  # 預設週三
    current_weekday = schedule_date.weekday()

    # 計算到下一個斷點日的天數
    if current_weekday <= breakpoint_weekday:
        days_to_breakpoint = breakpoint_weekday - current_weekday
    else:
        days_to_breakpoint = 7 - (current_weekday - breakpoint_weekday)

    week_end = schedule_date + timedelta(days=days_to_breakpoint)
    return week_end

def calculate_eta_target_date(week_end, eta_text):
    """
    根據 ETA 文字和斷點週結束日計算目標日期
    ETA 格式: 本週X, 下週X, 下下週X

    以斷點週結束日為基準:
    - 本週X = 斷點週結束日 + (X - 斷點週結束日的星期)
    - 下週X = 斷點週結束日 + 7 + (X - 斷點週結束日的星期)
    - 下下週X = 斷點週結束日 + 14 + (X - 斷點週結束日的星期)
    """
    # ETA 星期對應 (週一=0, 週二=1, ..., 週日=6)
    eta_weekday_map = {
        '一': 0, '二': 1, '三': 2, '四': 3,
        '五': 4, '六': 5, '日': 6, '天': 6,
    }

    if not eta_text or pd.isna(eta_text):
        return None

    eta_text = str(eta_text).strip()

    # 解析 ETA
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

    target_weekday = eta_weekday_map.get(weekday_char, 1)  # 預設週二
    week_end_weekday = week_end.weekday()

    # 計算目標日期
    # 從斷點週結束日到目標星期的天數差
    days_diff = target_weekday - week_end_weekday
    # 加上週數偏移
    target_date = week_end + timedelta(days=7 * weeks_offset + days_diff)

    return target_date

def find_week_column(target_date, date_columns):
    """
    找到目標日期所在週的欄位
    date_columns: {week_start_date: excel_col_idx}
    Forecast 日期已經是週一
    """
    # 轉換為 date 物件
    if isinstance(target_date, datetime):
        target_date = target_date.date()

    # 計算目標日期所在週的週一
    days_since_monday = target_date.weekday()
    week_monday = target_date - timedelta(days=days_since_monday)

    # 直接查找週一
    if week_monday in date_columns:
        return date_columns[week_monday]

    # 如果找不到，嘗試在範圍內查找
    for week_start, col_idx in date_columns.items():
        week_end = week_start + timedelta(days=6)
        if week_start <= target_date <= week_end:
            return col_idx

    return None

def test_erp_to_forecast():
    print("=" * 70)
    print("Test Pegatron ERP -> Forecast 邏輯 (使用 COM 保留格式)")
    print("=" * 70)

    # 檔案路徑
    input_forecast = r'd:\github\business_forecasting_pc\test\forecast_with_transit.xls'
    output_forecast = r'd:\github\business_forecasting_pc\test\forecast_with_erp.xls'
    erp_file = r'd:\github\business_forecasting_pc\test\integrated_erp.xlsx'

    # 如果 transit 輸出不存在，使用原始 forecast
    if not os.path.exists(input_forecast):
        input_forecast = r'd:\github\business_forecasting_pc\test\cleaned_forecast.xls'

    # 1. 先用 pandas 讀取資料分析結構
    forecast_df = pd.read_excel(input_forecast, header=None)
    print(f"Forecast 行數: {len(forecast_df)}, 欄數: {len(forecast_df.columns)}")

    erp_df = pd.read_excel(erp_file)
    print(f"ERP 行數: {len(erp_df)}")

    # 2. 建立 Forecast 區塊結構
    # F+G 欄位 = Line 客戶採購單號, I 欄位 (row+1) = Ordered Item
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

            forecast_row = row_idx + 1  # FORECAST 行 (Excel 1-based)

            forecast_blocks.append({
                'start_row': row_idx,
                'line_po': line_po,
                'ordered_item': ordered_item,
                'forecast_row': forecast_row + 1,  # Excel 是 1-based
            })
            row_idx += 8
        else:
            row_idx += 1

    print(f"\n找到 {len(forecast_blocks)} 個 Forecast 區塊:")
    for block in forecast_blocks:
        print(f"  Line PO={block['line_po']}, Ordered Item={block['ordered_item']}, FORECAST Excel Row={block['forecast_row']}")

    # 3. 取得日期欄位對應
    date_columns = {}
    for col_idx in range(14, len(forecast_df.columns)):
        date_val = forecast_df.iloc[1, col_idx]
        if pd.notna(date_val):
            if isinstance(date_val, str):
                try:
                    date_obj = pd.to_datetime(date_val)
                    date_columns[date_obj.date()] = col_idx + 1  # Excel 是 1-based
                except:
                    pass
            elif isinstance(date_val, (datetime, pd.Timestamp)):
                date_columns[date_val.date()] = col_idx + 1

    print(f"\n日期欄位對應: {len(date_columns)} 個日期")
    for d, col in list(date_columns.items())[:5]:
        print(f"  {d} -> Col {col}")

    # 4. 處理 ERP 資料，建立要更新的清單
    updates = []  # [(excel_row, excel_col, value), ...]

    # 過濾有客戶需求地區的 ERP
    erp_with_mapping = erp_df[erp_df['客戶需求地區'].notna() & (erp_df['客戶需求地區'] != '')]
    print(f"\n有 mapping 的 ERP 記錄: {len(erp_with_mapping)}")

    for idx, erp_row in erp_with_mapping.iterrows():
        erp_line_po = str(erp_row['Line 客戶採購單號']).strip() if pd.notna(erp_row['Line 客戶採購單號']) else ''
        erp_pn = str(erp_row['客戶料號']).strip() if pd.notna(erp_row['客戶料號']) else ''
        erp_qty = erp_row['淨需求'] if pd.notna(erp_row['淨需求']) else 0
        erp_schedule_date = erp_row['排程出貨日期']
        erp_breakpoint = str(erp_row['排程出貨日期斷點']).strip() if pd.notna(erp_row['排程出貨日期斷點']) else ''
        erp_eta = str(erp_row['ETA']).strip() if pd.notna(erp_row['ETA']) else ''

        if not erp_line_po or not erp_pn or erp_qty == 0:
            continue

        # 找到匹配的 Forecast 區塊
        matched_block = None
        for block in forecast_blocks:
            if block['line_po'] == erp_line_po and block['ordered_item'] == erp_pn:
                matched_block = block
                break

        if not matched_block:
            continue

        # 計算目標日期
        if pd.isna(erp_schedule_date) or not erp_breakpoint or not erp_eta:
            continue

        schedule_date = pd.to_datetime(erp_schedule_date)
        week_end = get_week_end_by_breakpoint(schedule_date, erp_breakpoint)
        target_date = calculate_eta_target_date(week_end, erp_eta)

        if target_date is None:
            continue

        # 找到對應的欄位
        excel_col = find_week_column(target_date, date_columns)

        if excel_col is None:
            print(f"  -> 警告: 找不到 ETA 日期 {target_date.date()} 對應的欄位")
            continue

        forecast_value = erp_qty * 1000

        print(f"\nERP: Line PO={erp_line_po}, PN={erp_pn}, Qty={erp_qty}")
        print(f"  排程={schedule_date.date()}, 斷點={erp_breakpoint}, ETA={erp_eta}")
        print(f"  -> 斷點週結束日={week_end.date()}")
        print(f"  -> 目標日期={target_date.date()} -> Excel Row {matched_block['forecast_row']}, Col {excel_col}, 值 = {forecast_value}")

        updates.append((matched_block['forecast_row'], excel_col, forecast_value))

    if not updates:
        print("\n沒有需要更新的 ERP 資料")
        return

    # 5. 使用 pywin32 COM 更新 Excel，保留格式和公式
    print(f"\n=== 使用 COM 更新 {len(updates)} 個儲存格 ===")

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
        for row, col, value in updates:
            key = (row, col)
            if key in update_dict:
                update_dict[key] += value
            else:
                update_dict[key] = value

        # 更新儲存格（只填數值，不動格式和公式）
        for (row, col), value in update_dict.items():
            current_val = ws.Cells(row, col).Value
            if current_val is None or current_val == '' or current_val == 0:
                ws.Cells(row, col).Value = value
            else:
                ws.Cells(row, col).Value = float(current_val) + value
            print(f"  更新 Row {row}, Col {col} = {ws.Cells(row, col).Value}")

        # 使用 SaveCopyAs 保存
        wb.SaveCopyAs(abs_path + ".tmp")
        wb.Close(SaveChanges=False)
        wb = None

        # 替換原檔案
        if os.path.exists(abs_path):
            os.remove(abs_path)
        os.rename(abs_path + ".tmp", abs_path)

        print(f"\n已輸出到: {output_forecast}")

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
    test_erp_to_forecast()
