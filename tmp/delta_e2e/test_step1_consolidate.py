"""
Delta Step 1 E2E Test: 8 檔案合併為固定方案二格式
"""
import os
import sys
import io

# force UTF-8 for Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from delta_forecast_processor import consolidate
import openpyxl

SRC_DIR = r"C:\Users\petty\Desktop\客戶相關資料\01.強茂\台達業務"
OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__)))

NBSP = '\xa0'
FORECAST_FILES = [
    f"PSBG{NBSP}PSB5-{NBSP}Ketwadee0406(完成).xlsx",
    f"PSBG{NBSP}PSB7_Kanyanat.S0406(完成).xlsx",
    f"PSBG{NBSP}PSB7-Weeraya0406(完成).xlsx",
    "PSBG (India IAI1&UPI2&DFI1 DIODES)-Jack0401.xlsx",
    "PSBG PSW1+CEW1- 楊洋彙整 0330(完成) (002).xlsx",
    "強茂 MWC1IPC1 MRP 03.30.2026.xlsx",
    "NBQ1.xlsx",
    "MRP(SVC1PWC1 DIODE&MOS)-100109-2026-3-30.xlsx",
]

REFERENCE = os.path.abspath(os.path.join(
    os.path.dirname(__file__), '..', '..', 'compare', 'delta', 'consolidated_template.xlsx'
))

OUTPUT = os.path.join(OUTPUT_DIR, 'step1_consolidated.xlsx')


def main():
    forecast_paths = [os.path.join(SRC_DIR, f) for f in FORECAST_FILES]

    for fp in forecast_paths:
        if not os.path.exists(fp):
            print(f"❌ 檔案不存在: {fp}")
            return 1

    print(f"📁 來源目錄: {SRC_DIR}")
    print(f"📄 Forecast 檔案數: {len(forecast_paths)}")
    print(f"📋 參考模板: {REFERENCE}")
    print(f"💾 輸出: {OUTPUT}")
    print()

    # PLANT codes 從 mapping 表抽取 (這裡用常見代碼模擬)
    plant_codes = [
        'PSB5', 'PSB7', 'IAI1', 'UPI2', 'DFI1',
        'PSW1', 'CEW1', 'MWC1', 'IPC1', 'NBQ1',
        'SVC1', 'PWC1'
    ]

    result = consolidate(
        forecast_files=forecast_paths,
        reference_template=REFERENCE,
        output_path=OUTPUT,
        plant_codes=plant_codes,
    )

    print()
    print("=" * 60)
    print(f"結果: {result.get('message', '')}")
    print(f"成功: {result.get('success', False)}")
    print(f"料號數: {result.get('part_count', 0)}")
    print("格式統計:")
    for k, v in result.get('format_stats', {}).items():
        print(f"  {k}: {v} 料號")
    print("=" * 60)

    if not result.get('success'):
        return 1

    # 驗證輸出檔案的固定 26 欄格式
    print()
    print("🔍 驗證輸出檔案結構...")
    wb = openpyxl.load_workbook(OUTPUT, read_only=True, data_only=True)
    ws = wb.active
    headers = []
    for col in range(1, 36):  # A~AI
        v = ws.cell(row=1, column=col).value
        headers.append((col, v))
    wb.close()

    print(f"  總欄位數 (A~AI): {len(headers)}")
    print(f"  固定欄位 A~I:")
    for col, v in headers[:9]:
        print(f"    col {col}: {v}")
    print(f"  日期欄位 J~AI:")
    for col, v in headers[9:]:
        print(f"    col {col}: {v}")

    # 驗證方案二固定格式
    date_headers = [v for (_, v) in headers[9:]]
    assert len(date_headers) == 26, f"期望 26 個日期欄位，實際 {len(date_headers)}"
    assert date_headers[0] == 'PASSDUE', f"col J 應為 PASSDUE, 實際 {date_headers[0]}"

    weekly = date_headers[1:17]
    monthly = date_headers[17:26]
    assert all(isinstance(w, str) and len(str(w)) == 8 and str(w).isdigit()
               for w in weekly), f"K~Z 應為 16 個 YYYYMMDD 週日期, 實際 {weekly}"
    MONTH_NAMES = ('JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
                   'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC')
    assert all(m in MONTH_NAMES for m in monthly), f"AA~AI 應為 9 個月份縮寫, 實際 {monthly}"
    assert len(monthly) == 9, f"AA~AI 應為 9 個月份, 實際 {len(monthly)}"

    print()
    print("✅ Step 1 通過 - 固定 26 欄格式正確")
    print(f"   PASSDUE | {weekly[0]}~{weekly[-1]} (16 週) | {monthly[0]}~{monthly[-1]} (9 月)")

    return 0


if __name__ == '__main__':
    sys.exit(main())
