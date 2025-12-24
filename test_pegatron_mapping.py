# -*- coding: utf-8 -*-
"""
Test Pegatron mapping logic
"""
import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Simulate customer_mappings data
# customer_name + region (M欄前4字) + delivery_location -> mapping values
sample_mappings = [
    {'customer_name': '和碩', 'region': '5699', 'delivery_location': '和碩-越南',
     'schedule_breakpoint': 'BP1', 'etd': '3', 'eta': '7'},
    {'customer_name': '和碩', 'region': '5691', 'delivery_location': '和碩-越南',
     'schedule_breakpoint': 'BP2', 'etd': '4', 'eta': '8'},
    {'customer_name': '和碩', 'region': '3A39', 'delivery_location': '和碩-新寧',
     'schedule_breakpoint': 'BP3', 'etd': '2', 'eta': '5'},
]

def test_erp_mapping():
    print("=" * 60)
    print("Test ERP Mapping Logic for Pegatron")
    print("=" * 60)

    # Read ERP file
    erp_df = pd.read_excel(r'd:\github\business_forecasting_pc\test\ERP.xlsx')
    print(f"ERP rows: {len(erp_df)}")

    # Key columns
    # D (idx 3): 客戶簡稱
    # M (idx 12): Line 客戶採購單號
    # AG (idx 32): 送貨地點
    customer_col = erp_df.columns[3]  # 客戶簡稱
    line_col = erp_df.columns[12]     # Line 客戶採購單號
    delivery_col = erp_df.columns[32] # 送貨地點

    print(f"Customer col: {customer_col}")
    print(f"Line col: {line_col}")
    print(f"Delivery col: {delivery_col}")

    # Build mapping lookup: (customer_name, region, delivery_location) -> values
    mapping_lookup = {}
    for m in sample_mappings:
        key = (m['customer_name'], m['region'], m['delivery_location'])
        mapping_lookup[key] = {
            'region': m['region'],
            'schedule_breakpoint': m['schedule_breakpoint'],
            'etd': m['etd'],
            'eta': m['eta']
        }

    print(f"\nMapping lookup keys: {list(mapping_lookup.keys())}")

    # Apply mapping
    def get_mapping_value(row, field):
        customer = str(row[customer_col]) if pd.notna(row[customer_col]) else ''
        line_po = str(row[line_col]) if pd.notna(row[line_col]) else ''
        delivery = str(row[delivery_col]) if pd.notna(row[delivery_col]) else ''

        # Get first 4 characters of Line 客戶採購單號 as region key
        region_key = line_po[:4] if len(line_po) >= 4 else line_po

        key = (customer, region_key, delivery)
        mapping = mapping_lookup.get(key, {})
        return mapping.get(field, '')

    erp_df['客戶需求地區'] = erp_df.apply(lambda row: get_mapping_value(row, 'region'), axis=1)
    erp_df['排程出貨日期斷點'] = erp_df.apply(lambda row: get_mapping_value(row, 'schedule_breakpoint'), axis=1)
    erp_df['ETD'] = erp_df.apply(lambda row: get_mapping_value(row, 'etd'), axis=1)
    erp_df['ETA'] = erp_df.apply(lambda row: get_mapping_value(row, 'eta'), axis=1)

    # Show results
    print("\n=== ERP Mapping Results (first 15 rows) ===")
    result_cols = [customer_col, line_col, delivery_col, '客戶需求地區', '排程出貨日期斷點', 'ETD', 'ETA']
    print(erp_df[result_cols].head(15).to_string())

    # Count matched
    matched = (erp_df['客戶需求地區'] != '').sum()
    print(f"\nMatched rows: {matched} / {len(erp_df)}")

    return erp_df

def test_transit_mapping(erp_df):
    print("\n" + "=" * 60)
    print("Test Transit Mapping Logic for Pegatron")
    print("=" * 60)

    # Read Transit file
    transit_df = pd.read_excel(r'd:\github\business_forecasting_pc\test\在途.xlsx')
    print(f"Transit rows: {len(transit_df)}")

    # Key columns
    # E (idx 4): Ordered Item -> matches ERP N (客戶料號)
    # L (idx 11): Line 客戶採購單號 -> matches ERP M
    ordered_item_col = transit_df.columns[4]  # Ordered Item
    line_col = transit_df.columns[11]         # Line 客戶採購單號

    print(f"Ordered Item col: {ordered_item_col}")
    print(f"Line col: {line_col}")

    # ERP columns for matching
    erp_line_col = erp_df.columns[12]    # M: Line 客戶採購單號
    erp_pn_col = erp_df.columns[13]      # N: 客戶料號

    # Build ERP lookup: (Line 客戶採購單號, 客戶料號) -> mapping values
    erp_lookup = {}
    for idx, row in erp_df.iterrows():
        line_po = str(row[erp_line_col]) if pd.notna(row[erp_line_col]) else ''
        pn = str(row[erp_pn_col]) if pd.notna(row[erp_pn_col]) else ''

        if line_po and pn:
            key = (line_po, pn)
            if key not in erp_lookup:  # Keep first match
                erp_lookup[key] = {
                    'region': row.get('客戶需求地區', ''),
                    'schedule_breakpoint': row.get('排程出貨日期斷點', ''),
                    'etd': row.get('ETD', ''),
                    'eta': row.get('ETA', '')
                }

    print(f"\nERP lookup entries: {len(erp_lookup)}")
    print(f"Sample keys: {list(erp_lookup.keys())[:5]}")

    # Apply mapping to transit
    def get_transit_mapping(row, field):
        line_po = str(row[line_col]) if pd.notna(row[line_col]) else ''
        ordered_item = str(row[ordered_item_col]) if pd.notna(row[ordered_item_col]) else ''

        key = (line_po, ordered_item)
        mapping = erp_lookup.get(key, {})
        return mapping.get(field, '')

    transit_df['客戶需求地區'] = transit_df.apply(lambda row: get_transit_mapping(row, 'region'), axis=1)
    transit_df['排程出貨日期斷點'] = transit_df.apply(lambda row: get_transit_mapping(row, 'schedule_breakpoint'), axis=1)
    transit_df['ETD'] = transit_df.apply(lambda row: get_transit_mapping(row, 'etd'), axis=1)
    transit_df['ETA_mapping'] = transit_df.apply(lambda row: get_transit_mapping(row, 'eta'), axis=1)

    # Show results
    print("\n=== Transit Mapping Results (first 15 rows) ===")
    result_cols = [ordered_item_col, line_col, '客戶需求地區', '排程出貨日期斷點', 'ETD', 'ETA_mapping']
    print(transit_df[result_cols].head(15).to_string())

    # Count matched
    matched = (transit_df['客戶需求地區'] != '').sum()
    print(f"\nMatched rows: {matched} / {len(transit_df)}")

if __name__ == "__main__":
    erp_df = test_erp_mapping()
    test_transit_mapping(erp_df)
