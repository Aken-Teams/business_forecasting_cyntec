# -*- coding: utf-8 -*-
"""
執行 Quanta Forecast 處理並輸出到 test/3 資料夾
"""
import sys
import os
import shutil

sys.stdout.reconfigure(encoding='utf-8')

# 切換到專案目錄
os.chdir(r'd:\github\business_forecasting_pc')

from ultra_fast_forecast_processor import UltraFastForecastProcessor

# 測試資料路徑
TEST_FOLDER = "test/3"
FORECAST_FILE = os.path.join(TEST_FOLDER, "cleaned_forecast.xlsx")
ERP_FILE = os.path.join(TEST_FOLDER, "integrated_erp.xlsx")
TRANSIT_FILE = os.path.join(TEST_FOLDER, "integrated_transit.xlsx")
OUTPUT_FOLDER = TEST_FOLDER
OUTPUT_FILENAME = "forecast_result.xlsx"

def main():
    print("=" * 70)
    print("執行 Quanta Forecast 處理")
    print("=" * 70)
    print(f"Forecast 檔案: {FORECAST_FILE}")
    print(f"ERP 檔案: {ERP_FILE}")
    print(f"Transit 檔案: {TRANSIT_FILE}")
    print(f"輸出資料夾: {OUTPUT_FOLDER}")
    print(f"輸出檔名: {OUTPUT_FILENAME}")
    print("=" * 70)

    # 先清除 ERP 和 Transit 的「已分配」狀態，模擬全新處理
    import pandas as pd
    import time

    print("\n🔄 清除已分配狀態，模擬全新處理...")

    # 使用複製檔案的方式避免鎖定問題
    import shutil

    erp_temp = os.path.join(TEST_FOLDER, "integrated_erp_temp.xlsx")
    transit_temp = os.path.join(TEST_FOLDER, "integrated_transit_temp.xlsx")

    try:
        erp_df = pd.read_excel(ERP_FILE)
        if '已分配' in erp_df.columns:
            erp_df['已分配'] = ''
        erp_df.to_excel(erp_temp, index=False)
        print(f"✅ 已清除 ERP 已分配狀態 (臨時檔案)")
    except PermissionError:
        print(f"⚠️ ERP 檔案被鎖定，跳過清除步驟")
        erp_temp = ERP_FILE

    try:
        transit_df = pd.read_excel(TRANSIT_FILE)
        if '已分配' in transit_df.columns:
            transit_df['已分配'] = ''
        transit_df.to_excel(transit_temp, index=False)
        print(f"✅ 已清除 Transit 已分配狀態 (臨時檔案)")
    except PermissionError:
        print(f"⚠️ Transit 檔案被鎖定，跳過清除步驟")
        transit_temp = TRANSIT_FILE

    # 使用臨時檔案作為輸入
    actual_erp = erp_temp
    actual_transit = transit_temp

    # 建立處理器 (使用臨時檔案)
    processor = UltraFastForecastProcessor(
        forecast_file=FORECAST_FILE,
        erp_file=actual_erp,
        transit_file=actual_transit,
        output_folder=OUTPUT_FOLDER,
        output_filename=OUTPUT_FILENAME
    )

    # 執行處理
    success = processor.process_all_blocks()

    # 確保輸出檔案存在（即使沒有修改）
    output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILENAME)
    if not os.path.exists(output_path):
        shutil.copy2(FORECAST_FILE, output_path)
        print(f"\n📁 已複製 Forecast 檔案到: {output_path}")

    if success:
        print("\n" + "=" * 70)
        print("處理完成！")
        print("=" * 70)
        print(f"輸出檔案: {output_path}")

        # 將臨時檔案的已分配狀態複製回原始檔案
        print("\n📋 同步已分配狀態到原始檔案...")
        try:
            if erp_temp != ERP_FILE and os.path.exists(erp_temp):
                shutil.copy2(erp_temp, ERP_FILE)
                print(f"✅ ERP 已分配狀態已同步到: {ERP_FILE}")
        except PermissionError:
            print(f"⚠️ 無法寫入 ERP 原始檔案（檔案被鎖定）")
        except Exception as e:
            print(f"⚠️ ERP 同步失敗: {e}")

        try:
            if transit_temp != TRANSIT_FILE and os.path.exists(transit_temp):
                shutil.copy2(transit_temp, TRANSIT_FILE)
                print(f"✅ Transit 已分配狀態已同步到: {TRANSIT_FILE}")
        except PermissionError:
            print(f"⚠️ 無法寫入 Transit 原始檔案（檔案被鎖定）")
        except Exception as e:
            print(f"⚠️ Transit 同步失敗: {e}")

        print(f"\nERP 已更新: {ERP_FILE}")
        print(f"Transit 已更新: {TRANSIT_FILE}")
    else:
        print("\n處理失敗！")

    # 清理臨時檔案
    for temp_file in [erp_temp, transit_temp]:
        if os.path.exists(temp_file) and temp_file != ERP_FILE and temp_file != TRANSIT_FILE:
            try:
                os.remove(temp_file)
            except:
                pass

    return success


if __name__ == "__main__":
    main()
