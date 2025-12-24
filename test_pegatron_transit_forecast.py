# -*- coding: utf-8 -*-
"""
Test Pegatron Transit -> Forecast 邏輯
將在途的 ETA QTY 填入 Forecast 對應日期欄位
使用 pywin32 COM 保留格式和公式
"""
import pandas as pd
import sys
import os
import shutil
import pythoncom
from win32com import client as win32
from datetime import datetime
sys.stdout.reconfigure(encoding='utf-8')

def test_transit_to_forecast():
    print("=" * 70)
    print("Test Pegatron Transit -> Forecast 邏輯 (使用 COM 保留格式)")
    print("=" * 70)

    # 檔案路徑
    input_forecast = r'd:\github\business_forecasting_pc\test\cleaned_forecast.xls'
    output_forecast = r'd:\github\business_forecasting_pc\test\forecast_with_transit.xls'
    transit_file = r'd:\github\business_forecasting_pc\test\在途.xlsx'

    # 1. 先用 pandas 讀取資料分析結構
    forecast_df = pd.read_excel(input_forecast, header=None)
    print(f"Forecast 行數: {len(forecast_df)}, 欄數: {len(forecast_df.columns)}")

    transit_df = pd.read_excel(transit_file)
    print(f"Transit 行數: {len(transit_df)}")

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

            eta_qty_row = row_idx + 4  # ETA QTY 行

            forecast_blocks.append({
                'start_row': row_idx,
                'line_po': line_po,
                'ordered_item': ordered_item,
                'eta_qty_row': eta_qty_row,
                'excel_row': eta_qty_row + 1  # Excel 是 1-based
            })
            row_idx += 8
        else:
            row_idx += 1

    print(f"\n找到 {len(forecast_blocks)} 個 Forecast 區塊:")
    for block in forecast_blocks:
        print(f"  Line PO={block['line_po']}, Ordered Item={block['ordered_item']}, ETA QTY Excel Row={block['excel_row']}")

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

    # 4. 處理在途資料，建立要更新的清單
    updates = []  # [(excel_row, excel_col, value), ...]

    transit_with_line = transit_df[transit_df['Line 客戶採購單號'].notna()]
    print(f"\n有 Line 客戶採購單號 的在途記錄: {len(transit_with_line)}")

    for idx, transit_row in transit_with_line.iterrows():
        transit_line_po = str(transit_row['Line 客戶採購單號']).strip()
        transit_ordered_item = str(transit_row['Ordered Item']).strip()
        transit_qty = transit_row['Qty'] if pd.notna(transit_row['Qty']) else 0
        transit_eta = transit_row['ETA']

        print(f"\n在途: Line PO={transit_line_po}, Ordered Item={transit_ordered_item}, Qty={transit_qty}, ETA={transit_eta}")

        matched_block = None
        for block in forecast_blocks:
            if block['line_po'] == transit_line_po and block['ordered_item'] == transit_ordered_item:
                matched_block = block
                break

        if matched_block:
            print(f"  -> 匹配到 Forecast 區塊")

            if pd.notna(transit_eta):
                eta_date = pd.to_datetime(transit_eta).date()
                days_since_monday = eta_date.weekday()
                week_start = eta_date - pd.Timedelta(days=days_since_monday)

                if week_start in date_columns:
                    excel_col = date_columns[week_start]
                    eta_qty_value = transit_qty * 1000

                    print(f"  -> ETA {eta_date} -> 週 {week_start} -> Excel Row {matched_block['excel_row']}, Col {excel_col}, 值 = {eta_qty_value}")

                    updates.append((matched_block['excel_row'], excel_col, eta_qty_value))
                else:
                    print(f"  -> 警告: 找不到 ETA 日期 {week_start} 對應的欄位")
            else:
                print(f"  -> 警告: 沒有 ETA 日期")
        else:
            print(f"  -> 未找到匹配的 Forecast 區塊")

    if not updates:
        print("\n沒有需要更新的資料")
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
    test_transit_to_forecast()
