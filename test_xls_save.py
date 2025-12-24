# -*- coding: utf-8 -*-
"""
Test .xls file save issue
"""
import os
import shutil
import pythoncom
from win32com import client as win32
import time
import sys

# Force UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

# Test file path
test_input = r"d:\github\business_forecasting_pc\uploads\5\20251224_151054\forecast_data.xls"
test_output = r"d:\github\business_forecasting_pc\test_output.xls"

def test_save_methods():
    print("=" * 50)
    print("Test .xls file save")
    print("=" * 50)

    if not os.path.exists(test_input):
        print(f"ERROR: Test file not found: {test_input}")
        return

    # Copy file
    print(f"1. Copy file to {test_output}")
    shutil.copy2(test_input, test_output)

    abs_path = os.path.abspath(test_output)
    print(f"   Absolute path: {abs_path}")

    # Initialize COM
    print("2. Initialize COM...")
    pythoncom.CoInitialize()

    excel = None
    wb = None

    try:
        # Start Excel
        print("3. Start Excel...")
        start = time.time()
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        print(f"   Time: {time.time() - start:.2f}s")

        # Open file
        print("4. Open file...")
        start = time.time()
        wb = excel.Workbooks.Open(abs_path)
        ws = wb.Sheets(1)
        print(f"   Time: {time.time() - start:.2f}s")

        # Get info
        used_range = ws.UsedRange
        max_row = used_range.Rows.Count
        max_col = used_range.Columns.Count
        print(f"   File size: {max_row} rows x {max_col} cols")

        # Make a small change
        print("5. Modify a cell...")
        ws.Cells(1, 1).Value = ws.Cells(1, 1).Value

        # Test method 1: wb.Save()
        print("6. Test wb.Save()...")
        start = time.time()
        try:
            wb.Save()
            print(f"   OK: wb.Save() success! Time: {time.time() - start:.2f}s")
        except Exception as e:
            print(f"   FAIL: wb.Save() failed: {e}")

        # Close
        print("7. Close workbook...")
        start = time.time()
        wb.Close(SaveChanges=False)
        wb = None
        print(f"   Time: {time.time() - start:.2f}s")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("8. Quit Excel...")
        start = time.time()
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
                del excel
        except Exception as e:
            print(f"   Warning: {e}")
        pythoncom.CoUninitialize()
        print(f"   Time: {time.time() - start:.2f}s")

    # Cleanup
    if os.path.exists(test_output):
        os.remove(test_output)
        print("9. Test file cleaned")

    print("=" * 50)
    print("Test completed")
    print("=" * 50)

def test_alternative_xlrd():
    """Test xlrd + xlwt method"""
    print("\n" + "=" * 50)
    print("Test xlrd + xlwt method")
    print("=" * 50)

    if not os.path.exists(test_input):
        print(f"ERROR: Test file not found: {test_input}")
        return

    try:
        import xlrd
        import xlwt
        from xlutils.copy import copy

        print("1. Open file with xlrd...")
        start = time.time()
        rb = xlrd.open_workbook(test_input, formatting_info=True)
        print(f"   Time: {time.time() - start:.2f}s")

        print("2. Copy workbook...")
        start = time.time()
        wb = copy(rb)
        print(f"   Time: {time.time() - start:.2f}s")

        print("3. Save file...")
        start = time.time()
        wb.save(test_output)
        print(f"   OK: Save success! Time: {time.time() - start:.2f}s")

        # Cleanup
        if os.path.exists(test_output):
            os.remove(test_output)
            print("4. Test file cleaned")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_save_methods()
    test_alternative_xlrd()
