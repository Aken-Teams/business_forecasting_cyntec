# -*- coding: utf-8 -*-
"""
測試 Quanta Forecast 分配邏輯
用於調試為什麼某些 ERP 記錄沒有被分配
"""
import pandas as pd
import sys
import os
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding='utf-8')

# 測試資料路徑
TEST_FOLDER = "test/3"
FORECAST_FILE = os.path.join(TEST_FOLDER, "cleaned_forecast.xlsx")
ERP_FILE = os.path.join(TEST_FOLDER, "integrated_erp.xlsx")
TRANSIT_FILE = os.path.join(TEST_FOLDER, "integrated_transit.xlsx")


def get_week_end_day(breakpoint):
    """根據斷點文字獲取星期幾的數字"""
    weekday_map = {
        '禮拜一': 0, '禮拜二': 1, '禮拜三': 2, '禮拜四': 3,
        '禮拜五': 4, '禮拜六': 5, '禮拜日': 6, '星期日': 6
    }
    return weekday_map.get(breakpoint, 3)  # 預設禮拜四


def get_eta_weekday(weekday_text):
    """獲取ETA中的星期幾"""
    weekday_map = {
        '一': 0, '二': 1, '三': 2, '四': 3,
        '五': 4, '六': 5, '日': 6, '天': 6
    }
    return weekday_map.get(weekday_text, 1)  # 預設星期二


def calculate_eta_date(eta, week_start, week_end):
    """根據ETA計算目標日期"""
    try:
        # 以week_end為基準，找到對應的標準周別（禮拜天~禮拜六）
        days_to_saturday = (5 - week_end.weekday()) % 7
        current_saturday = week_end + timedelta(days=days_to_saturday)
        current_sunday = current_saturday - timedelta(days=6)

        if '下下週' in eta:
            target_weekday = get_eta_weekday(eta.replace('下下週', ''))
            next_next_sunday = current_sunday + timedelta(days=14)
            days_to_target = (target_weekday - next_next_sunday.weekday()) % 7
            return next_next_sunday + timedelta(days=days_to_target)
        elif '下週' in eta:
            target_weekday = get_eta_weekday(eta.replace('下週', ''))
            next_sunday = current_sunday + timedelta(days=7)
            days_to_target = (target_weekday - next_sunday.weekday()) % 7
            return next_sunday + timedelta(days=days_to_target)
        elif '本週' in eta:
            target_weekday = get_eta_weekday(eta.replace('本週', ''))
            days_to_target = (target_weekday - current_sunday.weekday()) % 7
            return current_sunday + timedelta(days=days_to_target)
        else:
            return None
    except Exception as e:
        print(f"    ❌ ETA計算失敗: {e}")
        return None


def normalize_date(date_value):
    """統一日期處理函數"""
    try:
        if pd.isna(date_value) or date_value is None:
            return None
        if isinstance(date_value, (datetime, pd.Timestamp)):
            return date_value
        if isinstance(date_value, str):
            date_str = str(date_value).strip()
            if not date_str or date_str.lower() in ['nan', 'none', '']:
                return None
            date_formats = [
                "%Y/%m/%d", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y%m%d",
            ]
            for fmt in date_formats:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
            try:
                return pd.to_datetime(date_str)
            except:
                return None
        try:
            return pd.to_datetime(date_value)
        except:
            return None
    except Exception as e:
        return None


def find_target_position(forecast_df, target_date_str, start_row, end_row):
    """找到目標填寫位置"""
    try:
        for col_idx in range(10, min(49, len(forecast_df.columns))):
            start_date_row = start_row + 1
            end_date_row = start_row + 2

            if start_date_row < len(forecast_df) and end_date_row < len(forecast_df):
                try:
                    start_date = forecast_df.iloc[start_date_row, col_idx]
                    end_date = forecast_df.iloc[end_date_row, col_idx]

                    if pd.notna(start_date) and pd.notna(end_date):
                        start_date_str = str(int(float(start_date)))
                        end_date_str = str(int(float(end_date)))

                        if len(start_date_str) == 8 and len(end_date_str) == 8:
                            if start_date_str <= target_date_str <= end_date_str:
                                # 找到目標欄位，現在找供應數量行
                                for row_idx in range(start_row, min(start_row + 18, len(forecast_df))):
                                    k_value = forecast_df.iloc[row_idx, 10]  # K欄位
                                    if pd.notna(k_value) and str(k_value) == "供應數量":
                                        return col_idx, row_idx + 2
                except (ValueError, TypeError):
                    continue
        return None, None
    except Exception as e:
        return None, None


def test_allocation_debug():
    print("=" * 70)
    print("Quanta Forecast 分配調試測試")
    print("=" * 70)

    # 1. 讀取資料
    print("\n=== 1. 讀取資料 ===")
    forecast_df = pd.read_excel(FORECAST_FILE)
    erp_df = pd.read_excel(ERP_FILE)

    print(f"Forecast 行數: {len(forecast_df)}, 欄數: {len(forecast_df.columns)}")
    print(f"ERP 行數: {len(erp_df)}")

    # 檢查 ERP 必要欄位
    required_cols = ['客戶料號', '客戶需求地區', '淨需求', '排程出貨日期', '排程出貨日期斷點', 'ETA', '已分配']
    missing_cols = [c for c in required_cols if c not in erp_df.columns]
    if missing_cols:
        print(f"⚠️ ERP 缺少欄位: {missing_cols}")

    # 2. 篩選有 mapping 的 ERP 記錄
    print("\n=== 2. 篩選有 mapping 的 ERP 記錄 ===")
    erp_with_mapping = erp_df[erp_df['客戶需求地區'].notna() & (erp_df['客戶需求地區'] != '')]
    print(f"有 mapping 的 ERP 記錄: {len(erp_with_mapping)}")

    # 統計已分配狀態
    if '已分配' in erp_df.columns:
        allocated_count = (erp_df['已分配'] == '✓').sum()
        print(f"已分配筆數: {allocated_count}")
        not_allocated = erp_with_mapping[erp_with_mapping['已分配'] != '✓']
        print(f"有 mapping 但未分配: {len(not_allocated)}")

    # 3. 找到 Forecast 數據塊
    print("\n=== 3. 識別 Forecast 數據塊 ===")
    data_blocks = []
    a_col = forecast_df.columns[0]  # A欄位
    d_col = forecast_df.columns[3]  # D欄位

    current_block = None
    for idx, row in forecast_df.iterrows():
        customer_part = row[a_col]
        customer_region = row[d_col]

        if pd.notna(customer_part) and pd.notna(customer_region) and \
           customer_part != "需求週數" and customer_part != "客戶料號" and \
           len(str(customer_part)) > 5:

            if current_block is None or \
               current_block['customer_part'] != customer_part or \
               current_block['customer_region'] != customer_region:

                if current_block is not None:
                    data_blocks.append(current_block)

                current_block = {
                    'customer_part': customer_part,
                    'customer_region': customer_region,
                    'start_row': idx,
                    'end_row': idx
                }
            else:
                current_block['end_row'] = idx

    if current_block is not None:
        data_blocks.append(current_block)

    print(f"找到 {len(data_blocks)} 個數據塊")

    # 建立數據塊索引
    block_keys = set()
    for block in data_blocks:
        key = f"{block['customer_part']}_{block['customer_region']}"
        block_keys.add(key)

    # 4. 分析未分配的 ERP 記錄
    print("\n=== 4. 分析未分配的 ERP 記錄 ===")

    skip_reasons = {
        'already_allocated': 0,
        'no_forecast_block': 0,
        'no_schedule_date': 0,
        'no_breakpoint': 0,
        'no_eta': 0,
        'invalid_eta': 0,
        'no_target_column': 0,
        'success': 0,
    }

    for idx, erp_row in erp_with_mapping.iterrows():
        customer_part = str(erp_row['客戶料號']).strip() if pd.notna(erp_row['客戶料號']) else ''
        customer_region = str(erp_row['客戶需求地區']).strip() if pd.notna(erp_row['客戶需求地區']) else ''
        schedule_date = erp_row['排程出貨日期']
        breakpoint_text = str(erp_row['排程出貨日期斷點']).strip() if pd.notna(erp_row['排程出貨日期斷點']) else ''
        eta_text = str(erp_row['ETA']).strip() if pd.notna(erp_row['ETA']) else ''
        net_demand = erp_row['淨需求'] if pd.notna(erp_row['淨需求']) else 0
        already_allocated = erp_row.get('已分配', '') == '✓'

        # 跳過已分配的
        if already_allocated:
            skip_reasons['already_allocated'] += 1
            continue

        match_key = f"{customer_part}_{customer_region}"

        # 檢查是否有對應的 Forecast 區塊
        if match_key not in block_keys:
            skip_reasons['no_forecast_block'] += 1
            # 顯示找不到區塊的記錄
            print(f"\n⚠️ 找不到 Forecast 區塊:")
            print(f"   客戶料號: {customer_part}")
            print(f"   客戶需求地區: {customer_region}")
            print(f"   match_key: {match_key}")
            continue

        # 檢查排程日期
        if pd.isna(schedule_date):
            skip_reasons['no_schedule_date'] += 1
            print(f"\n⚠️ 排程日期為空:")
            print(f"   客戶料號: {customer_part}, 地區: {customer_region}")
            continue

        # 檢查斷點
        if not breakpoint_text:
            skip_reasons['no_breakpoint'] += 1
            print(f"\n⚠️ 斷點為空:")
            print(f"   客戶料號: {customer_part}, 地區: {customer_region}")
            continue

        # 檢查 ETA
        if not eta_text or eta_text.lower() in ['nan', 'none']:
            skip_reasons['no_eta'] += 1
            print(f"\n⚠️ ETA 為空:")
            print(f"   客戶料號: {customer_part}, 地區: {customer_region}")
            continue

        # 計算目標日期
        date_obj = normalize_date(schedule_date)
        if date_obj is None:
            skip_reasons['no_schedule_date'] += 1
            print(f"\n⚠️ 無法解析排程日期: {schedule_date}")
            continue

        week_end_day = get_week_end_day(breakpoint_text)
        days_to_week_end = (week_end_day - date_obj.weekday()) % 7
        week_end_date = date_obj + timedelta(days=days_to_week_end)
        week_start_date = week_end_date - timedelta(days=6)

        target_date = calculate_eta_date(eta_text, week_start_date, week_end_date)

        if target_date is None:
            skip_reasons['invalid_eta'] += 1
            print(f"\n⚠️ 無法計算 ETA 目標日期:")
            print(f"   客戶料號: {customer_part}, 地區: {customer_region}")
            print(f"   ETA: {eta_text}")
            continue

        target_date_str = target_date.strftime("%Y%m%d")

        # 找到對應的區塊
        matched_block = None
        for block in data_blocks:
            if block['customer_part'] == customer_part and block['customer_region'] == customer_region:
                matched_block = block
                break

        if matched_block is None:
            skip_reasons['no_forecast_block'] += 1
            continue

        # 找到目標位置
        col_idx, row_idx = find_target_position(forecast_df, target_date_str, matched_block['start_row'], matched_block['end_row'])

        if col_idx is None:
            skip_reasons['no_target_column'] += 1
            print(f"\n⚠️ 找不到目標欄位:")
            print(f"   客戶料號: {customer_part}, 地區: {customer_region}")
            print(f"   目標日期: {target_date_str} ({target_date.strftime('%Y-%m-%d')})")
            print(f"   排程日期: {date_obj.strftime('%Y-%m-%d')}, 斷點: {breakpoint_text}, ETA: {eta_text}")

            # 顯示 Forecast 區塊的日期範圍
            print(f"   Forecast 區塊 start_row: {matched_block['start_row']}")
            for col_i in range(10, min(20, len(forecast_df.columns))):
                try:
                    start_date = forecast_df.iloc[matched_block['start_row'] + 1, col_i]
                    end_date = forecast_df.iloc[matched_block['start_row'] + 2, col_i]
                    if pd.notna(start_date) and pd.notna(end_date):
                        print(f"      Col {col_i}: {start_date} ~ {end_date}")
                except:
                    pass
            continue

        skip_reasons['success'] += 1

    # 5. 輸出統計
    print("\n" + "=" * 70)
    print("分配統計結果")
    print("=" * 70)
    print(f"已分配 (跳過): {skip_reasons['already_allocated']}")
    print(f"找不到 Forecast 區塊: {skip_reasons['no_forecast_block']}")
    print(f"排程日期為空: {skip_reasons['no_schedule_date']}")
    print(f"斷點為空: {skip_reasons['no_breakpoint']}")
    print(f"ETA 為空: {skip_reasons['no_eta']}")
    print(f"無法計算 ETA: {skip_reasons['invalid_eta']}")
    print(f"找不到目標欄位: {skip_reasons['no_target_column']}")
    print(f"可成功分配: {skip_reasons['success']}")
    print(f"\n總計: {sum(skip_reasons.values())}")


if __name__ == "__main__":
    test_allocation_debug()
