# -*- coding: utf-8 -*-
"""
Test .xls file save with SaveAs
"""
import os
import shutil
import pythoncom
from win32com import client as win32
import time
import sys

sys.stdout.reconfigure(encoding='utf-8')

test_input = r"d:\github\business_forecasting_pc\uploads\5\20251224_151054\forecast_data.xls"
test_output = r"d:\github\business_forecasting_pc\test_output.xls"

def test_saveas():
    print("=" * 50)
    print("Test SaveAs method")
    print("=" * 50)

    if not os.path.exists(test_input):
        print(f"ERROR: Test file not found: {test_input}")
        return

    # Copy file
    print(f"1. Copy file to {test_output}")
    shutil.copy2(test_input, test_output)

    abs_path = os.path.abspath(test_output)
    print(f"   Absolute path: {abs_path}")

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        print("2. Start Excel...")
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        print("3. Open file...")
        wb = excel.Workbooks.Open(abs_path)
        ws = wb.Sheets(1)

        used_range = ws.UsedRange
        max_row = used_range.Rows.Count
        max_col = used_range.Columns.Count
        print(f"   File size: {max_row} rows x {max_col} cols")

        # Modify a cell
        print("4. Modify a cell...")
        ws.Cells(1, 1).Value = ws.Cells(1, 1).Value

        # Test SaveAs with xlExcel8 format (56)
        print("5. Test SaveAs with format 56 (xlExcel8)...")
        temp_output = abs_path + ".tmp.xls"
        start = time.time()
        try:
            wb.SaveAs(temp_output, FileFormat=56)
            print(f"   OK: SaveAs success! Time: {time.time() - start:.2f}s")
            # Close without saving again
            wb.Close(SaveChanges=False)
            wb = None
            # Replace original file
            if os.path.exists(abs_path):
                os.remove(abs_path)
            os.rename(temp_output, abs_path)
            print("   File replaced successfully")
        except Exception as e:
            print(f"   FAIL: SaveAs failed: {e}")
            if wb:
                wb.Close(SaveChanges=False)
                wb = None

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("6. Quit Excel...")
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
                del excel
        except Exception as e:
            print(f"   Warning: {e}")
        pythoncom.CoUninitialize()

    # Cleanup
    for f in [test_output, test_output + ".tmp.xls"]:
        if os.path.exists(f):
            os.remove(f)
    print("7. Test files cleaned")
    print("=" * 50)

if __name__ == "__main__":
    test_saveas()
