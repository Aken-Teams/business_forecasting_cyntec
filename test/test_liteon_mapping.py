"""測試光寶 ERP Mapping 邏輯"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from database import get_customer_mappings_raw

# 1. 讀取 ERP
erp_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'uploads', '6', '20260311_181617', 'erp_data.xlsx')
erp_df = pd.read_excel(erp_file)
print(f'ERP 行數: {len(erp_df)}')
print(f'ERP 欄位數: {len(erp_df.columns)}')

# 2. 建立 mapping lookup
mapping_records = get_customer_mappings_raw(6)
print(f'Mapping 記錄數: {len(mapping_records)}')

liteon_lookup_11 = {}
liteon_lookup_32 = {}
for m in mapping_records:
    cname = str(m['customer_name']).strip() if m['customer_name'] else ''
    order_type = str(m.get('order_type', '')).strip()
    delivery_loc = str(m.get('delivery_location', '')).strip() if m.get('delivery_location') else ''
    warehouse = str(m.get('warehouse', '')).strip() if m.get('warehouse') else ''

    mapping_values = {
        'region': str(m['region']).strip() if m['region'] else '',
        'schedule_breakpoint': str(m['schedule_breakpoint']).strip() if m['schedule_breakpoint'] else '',
        'etd': str(m['etd']).strip() if m['etd'] else '',
        'eta': str(m['eta']).strip() if m['eta'] else '',
        'date_calc_type': str(m.get('date_calc_type', '')).strip() if m.get('date_calc_type') else ''
    }

    if order_type == '11' and delivery_loc:
        liteon_lookup_11[(cname, delivery_loc)] = mapping_values
    elif order_type == '32' and warehouse:
        liteon_lookup_32[(cname, warehouse)] = mapping_values

print(f'Lookup 11: {len(liteon_lookup_11)} 筆')
print(f'Lookup 32: {len(liteon_lookup_32)} 筆')

# 3. 欄位名稱
customer_col = '客戶簡稱'
delivery_col = '送貨地點'
warehouse_col = '倉庫'
order_type_col = '訂單型態'

# 4. 執行 mapping
def get_liteon_mapping(row, field):
    customer = str(row[customer_col]).strip() if pd.notna(row[customer_col]) else ''
    order_type_val = str(row[order_type_col]).strip() if pd.notna(row[order_type_col]) else ''
    ot_prefix = order_type_val[:2] if len(order_type_val) >= 2 else order_type_val

    if ot_prefix == '11':
        delivery = str(row[delivery_col]).strip() if pd.notna(row[delivery_col]) else ''
        key = (customer, delivery)
        mapping = liteon_lookup_11.get(key, {})
    elif ot_prefix == '32':
        wh = str(row[warehouse_col]).strip() if pd.notna(row[warehouse_col]) else ''
        key = (customer, wh)
        mapping = liteon_lookup_32.get(key, {})
    else:
        mapping = {}
    return mapping.get(field, '') if mapping else ''

erp_df['客戶需求地區'] = erp_df.apply(lambda row: get_liteon_mapping(row, 'region'), axis=1)
erp_df['排程出貨日期斷點'] = erp_df.apply(lambda row: get_liteon_mapping(row, 'schedule_breakpoint'), axis=1)
erp_df['ETD_mapping'] = erp_df.apply(lambda row: get_liteon_mapping(row, 'etd'), axis=1)
erp_df['ETA_mapping'] = erp_df.apply(lambda row: get_liteon_mapping(row, 'eta'), axis=1)
erp_df['日期算法'] = erp_df.apply(lambda row: get_liteon_mapping(row, 'date_calc_type'), axis=1)

# 5. 統計
total = len(erp_df)
matched = (erp_df['客戶需求地區'] != '').sum()
unmatched = total - matched
print(f'\n=== 匹配結果 ===')
print(f'總行數: {total}')
print(f'匹配成功: {matched} ({matched/total*100:.1f}%)')
print(f'未匹配: {unmatched} ({unmatched/total*100:.1f}%)')

# 6. 未匹配分析
if unmatched > 0:
    miss_df = erp_df[erp_df['客戶需求地區'] == '']
    print(f'\n=== 未匹配記錄分析 ===')
    miss_combos = miss_df.groupby([customer_col, order_type_col, delivery_col, warehouse_col]).size().reset_index(name='count')
    for _, r in miss_combos.iterrows():
        ot = str(r[order_type_col])[:2]
        if ot == '11':
            print(f'  type={ot} | name={r[customer_col]} | delivery={r[delivery_col]} | count={r["count"]}')
        else:
            print(f'  type={ot} | name={r[customer_col]} | warehouse={r[warehouse_col]} | count={r["count"]}')

# 7. 匹配成功 sample
print(f'\n=== 匹配成功 sample (前10行) ===')
matched_df = erp_df[erp_df['客戶需求地區'] != ''].head(10)
for _, r in matched_df.iterrows():
    print(f'  {r[customer_col]} | {r[order_type_col]} | region={r["客戶需求地區"]} | etd={r["ETD_mapping"]} | eta={r["ETA_mapping"]} | calc={r["日期算法"]}')

# 8. 輸出
output_file = os.path.join(os.path.dirname(__file__), 'liteon_integrated_erp.xlsx')
erp_df.to_excel(output_file, index=False)
print(f'\n結果已輸出到 test/liteon_integrated_erp.xlsx')
print(f'輸出欄位數: {len(erp_df.columns)}')
print(f'新增欄位: 客戶需求地區, 排程出貨日期斷點, ETD_mapping, ETA_mapping, 日期算法')
