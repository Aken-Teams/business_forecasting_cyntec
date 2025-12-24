# -*- coding: utf-8 -*-
"""
Test Pegatron mapping logic v2 - 模擬實際 app.py 中的邏輯
使用真實資料庫的 mapping 資料
"""
import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 從資料庫讀取真實的 mapping 資料
from database import get_customer_mappings_raw

# Pegatron 的 user_id
PEGATRON_USER_ID = 5

def test_erp_mapping():
    print("=" * 60)
    print("Test ERP Mapping Logic for Pegatron (v2)")
    print("=" * 60)

    # Read ERP file
    erp_df = pd.read_excel(r'd:\github\business_forecasting_pc\test\ERP.xlsx')
    print(f"ERP rows: {len(erp_df)}")

    # 從資料庫讀取真實的 mapping 資料
    mapping_records = get_customer_mappings_raw(PEGATRON_USER_ID)
    print(f"從資料庫讀取 {len(mapping_records)} 筆 mapping 記錄")

    # 模擬 app.py 中的邏輯
    # 找客戶簡稱欄位
    customer_col = None
    for col in erp_df.columns:
        if '客戶' in str(col) and '簡稱' in str(col):
            customer_col = col
            break

    print(f"Customer col: {customer_col}")

    # 建立兩種 lookup:
    # 1. 完整匹配: (customer_name, region, delivery_location) -> mapping values
    # 2. 簡化匹配: (customer_name, region) -> mapping values (當 delivery_location 為空時)
    pegatron_mapping_lookup = {}  # 完整 3 欄位匹配
    pegatron_mapping_lookup_simple = {}  # 簡化 2 欄位匹配（不需要送貨地點）
    for m in mapping_records:
        customer_name = str(m['customer_name']).strip() if m['customer_name'] else ''
        region = str(m['region']).strip() if m['region'] else ''
        delivery_location = str(m['delivery_location']).strip() if m['delivery_location'] else ''

        mapping_values = {
            'region': region,
            'schedule_breakpoint': str(m['schedule_breakpoint']).strip() if m['schedule_breakpoint'] else '',
            'etd': str(m['etd']).strip() if m['etd'] else '',
            'eta': str(m['eta']).strip() if m['eta'] else ''
        }

        if delivery_location:
            # 有送貨地點的用完整 3 欄位匹配
            key = (customer_name, region, delivery_location)
            pegatron_mapping_lookup[key] = mapping_values
        else:
            # 沒有送貨地點的用簡化 2 欄位匹配
            key = (customer_name, region)
            pegatron_mapping_lookup_simple[key] = mapping_values

    print(f"完整匹配 lookup ({len(pegatron_mapping_lookup)}):")
    for key in pegatron_mapping_lookup.keys():
        print(f"  {key}")
    print(f"簡化匹配 lookup ({len(pegatron_mapping_lookup_simple)}):")
    for key in pegatron_mapping_lookup_simple.keys():
        print(f"  {key}")

    # 找到必要欄位
    line_po_col = erp_df.columns[12] if len(erp_df.columns) > 12 else None
    delivery_col = erp_df.columns[32] if len(erp_df.columns) > 32 else None

    print(f"Line PO col (M): {line_po_col}")
    print(f"Delivery col (AG): {delivery_col}")

    # 應用 Pegatron 映射
    def get_pegatron_mapping(row, field):
        customer = str(row[customer_col]).strip() if pd.notna(row[customer_col]) else ''
        line_po = str(row[line_po_col]).strip() if pd.notna(row[line_po_col]) else ''
        delivery = str(row[delivery_col]).strip() if pd.notna(row[delivery_col]) else ''

        # 取 Line 客戶採購單號的前 4 字作為 region key
        region_key = line_po[:4] if len(line_po) >= 4 else line_po

        # 先嘗試完整 3 欄位匹配
        key_full = (customer, region_key, delivery)
        mapping = pegatron_mapping_lookup.get(key_full)

        # 如果完整匹配失敗，嘗試簡化 2 欄位匹配（不需要送貨地點）
        if not mapping:
            key_simple = (customer, region_key)
            mapping = pegatron_mapping_lookup_simple.get(key_simple, {})

        return mapping.get(field, '') if mapping else ''

    erp_df['客戶需求地區'] = erp_df.apply(lambda row: get_pegatron_mapping(row, 'region'), axis=1)
    erp_df['排程出貨日期斷點'] = erp_df.apply(lambda row: get_pegatron_mapping(row, 'schedule_breakpoint'), axis=1)
    erp_df['ETD'] = erp_df.apply(lambda row: get_pegatron_mapping(row, 'etd'), axis=1)
    erp_df['ETA'] = erp_df.apply(lambda row: get_pegatron_mapping(row, 'eta'), axis=1)

    # Show results
    print("\n=== ERP Mapping Results (first 15 rows) ===")
    result_cols = [customer_col, line_po_col, delivery_col, '客戶需求地區', '排程出貨日期斷點', 'ETD', 'ETA']
    print(erp_df[result_cols].head(15).to_string())

    # Count matched
    matched = (erp_df['客戶需求地區'] != '').sum()
    print(f"\nMatched rows: {matched} / {len(erp_df)}")

    # 輸出到 test 資料夾
    output_path = r'd:\github\business_forecasting_pc\test\integrated_erp.xlsx'
    erp_df.to_excel(output_path, index=False)
    print(f"\n已輸出到: {output_path}")

    return erp_df

def test_transit_mapping(erp_df):
    print("\n" + "=" * 60)
    print("Test Transit Mapping Logic for Pegatron (v2)")
    print("=" * 60)

    # Read Transit file
    transit_df = pd.read_excel(r'd:\github\business_forecasting_pc\test\在途.xlsx')
    print(f"Transit rows: {len(transit_df)}")

    # Transit 欄位
    transit_ordered_item_col = transit_df.columns[4] if len(transit_df.columns) > 4 else None
    transit_line_po_col = transit_df.columns[11] if len(transit_df.columns) > 11 else None

    print(f"Transit Ordered Item col (E): {transit_ordered_item_col}")
    print(f"Transit Line PO col (L): {transit_line_po_col}")

    # ERP 欄位
    erp_line_po_col = erp_df.columns[12] if len(erp_df.columns) > 12 else None
    erp_pn_col = erp_df.columns[13] if len(erp_df.columns) > 13 else None

    print(f"ERP Line PO col (M): {erp_line_po_col}")
    print(f"ERP PN col (N): {erp_pn_col}")

    # 建立 ERP lookup: (Line 客戶採購單號, 客戶料號) -> mapping values
    erp_lookup = {}
    for idx, row in erp_df.iterrows():
        line_po = str(row[erp_line_po_col]).strip() if pd.notna(row[erp_line_po_col]) else ''
        pn = str(row[erp_pn_col]).strip() if pd.notna(row[erp_pn_col]) else ''

        if line_po and pn:
            key = (line_po, pn)
            if key not in erp_lookup:  # 保留第一筆匹配
                erp_lookup[key] = {
                    'region': str(row.get('客戶需求地區', '')).strip() if pd.notna(row.get('客戶需求地區', '')) else '',
                    'schedule_breakpoint': str(row.get('排程出貨日期斷點', '')).strip() if pd.notna(row.get('排程出貨日期斷點', '')) else '',
                    'etd': str(row.get('ETD', '')).strip() if pd.notna(row.get('ETD', '')) else '',
                    'eta': str(row.get('ETA', '')).strip() if pd.notna(row.get('ETA', '')) else ''
                }

    print(f"\nERP lookup entries: {len(erp_lookup)}")
    print(f"Sample keys: {list(erp_lookup.keys())[:5]}")

    # 應用 Transit 映射
    def get_pegatron_transit_mapping(row, field):
        line_po = str(row[transit_line_po_col]).strip() if pd.notna(row[transit_line_po_col]) else ''
        ordered_item = str(row[transit_ordered_item_col]).strip() if pd.notna(row[transit_ordered_item_col]) else ''

        key = (line_po, ordered_item)
        mapping = erp_lookup.get(key, {})
        return mapping.get(field, '')

    transit_df['客戶需求地區'] = transit_df.apply(lambda row: get_pegatron_transit_mapping(row, 'region'), axis=1)
    transit_df['排程出貨日期斷點'] = transit_df.apply(lambda row: get_pegatron_transit_mapping(row, 'schedule_breakpoint'), axis=1)
    transit_df['ETD'] = transit_df.apply(lambda row: get_pegatron_transit_mapping(row, 'etd'), axis=1)
    transit_df['ETA_mapping'] = transit_df.apply(lambda row: get_pegatron_transit_mapping(row, 'eta'), axis=1)

    # Show results
    print("\n=== Transit Mapping Results (all rows with data) ===")
    result_cols = [transit_ordered_item_col, transit_line_po_col, '客戶需求地區', '排程出貨日期斷點', 'ETD', 'ETA_mapping']

    # 顯示有 Line PO 的行
    has_line_po = transit_df[transit_line_po_col].notna()
    print(transit_df.loc[has_line_po, result_cols].to_string())

    # Count matched
    matched = (transit_df['客戶需求地區'] != '').sum()
    print(f"\nMatched rows: {matched} / {len(transit_df)}")

    # 顯示匹配的行
    print("\n=== Matched Transit rows ===")
    matched_rows = transit_df[transit_df['客戶需求地區'] != '']
    if len(matched_rows) > 0:
        print(matched_rows[result_cols].to_string())
    else:
        print("No matched rows")

    # 輸出到 test 資料夾
    output_path = r'd:\github\business_forecasting_pc\test\integrated_transit.xlsx'
    transit_df.to_excel(output_path, index=False)
    print(f"\n已輸出到: {output_path}")

if __name__ == "__main__":
    erp_df = test_erp_mapping()
    test_transit_mapping(erp_df)
