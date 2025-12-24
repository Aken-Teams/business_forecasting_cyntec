# -*- coding: utf-8 -*-
"""
Test .xls file save - check protection and try different methods
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

def test_with_unprotect():
    print("=" * 50)
    print("Test with Unprotect")
    print("=" * 50)

    if not os.path.exists(test_input):
        print(f"ERROR: Test file not found: {test_input}")
        return

    shutil.copy2(test_input, test_output)
    abs_path = os.path.abspath(test_output)

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        print("1. Start Excel...")
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        print("2. Open file...")
        wb = excel.Workbooks.Open(abs_path)
        ws = wb.Sheets(1)

        # Check workbook protection
        print("3. Check protection...")
        print(f"   Workbook.ProtectStructure: {wb.ProtectStructure}")
        print(f"   Workbook.ProtectWindows: {wb.ProtectWindows}")
        print(f"   Sheet.ProtectContents: {ws.ProtectContents}")
        print(f"   Sheet.ProtectDrawingObjects: {ws.ProtectDrawingObjects}")
        print(f"   Sheet.ProtectScenarios: {ws.ProtectScenarios}")

        # Try to unprotect
        print("4. Try to unprotect sheet...")
        try:
            ws.Unprotect()
            print("   Sheet unprotected")
        except Exception as e:
            print(f"   Unprotect failed: {e}")

        # Check if file is read-only
        print(f"5. Workbook.ReadOnly: {wb.ReadOnly}")

        # Modify a cell
        print("6. Modify a cell...")
        old_val = ws.Cells(1, 1).Value
        ws.Cells(1, 1).Value = old_val
        print(f"   Cell modified")

        # Try different save methods
        print("7. Try wb.Save()...")
        try:
            wb.Save()
            print("   OK!")
        except Exception as e:
            print(f"   FAIL: {e}")

        print("8. Try SaveCopyAs...")
        temp_copy = abs_path + ".copy.xls"
        try:
            wb.SaveCopyAs(temp_copy)
            print(f"   OK! Saved to {temp_copy}")
            if os.path.exists(temp_copy):
                os.remove(temp_copy)
        except Exception as e:
            print(f"   FAIL: {e}")

        wb.Close(SaveChanges=False)
        wb = None

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("9. Quit Excel...")
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

    # Cleanup
    for f in [test_output, test_output + ".copy.xls"]:
        if os.path.exists(f):
            os.remove(f)
    print("Done")

if __name__ == "__main__":
    test_with_unprotect()
