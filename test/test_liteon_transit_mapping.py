"""測試光寶 Transit Mapping 邏輯"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from database import get_customer_mappings_raw

# 1. 讀取檔案
base = os.path.dirname(os.path.dirname(__file__))
erp_df = pd.read_excel(os.path.join(base, 'uploads', '6', '20260311_181617', 'erp_data.xlsx'))
transit_df = pd.read_excel(os.path.join(base, 'uploads', '6', '20260311_181617', 'transit_data.xlsx'))
print(f'ERP: {len(erp_df)} rows, Transit: {len(transit_df)} rows')

# 2. 建立 mapping lookup (同 ERP mapping 邏輯)
mapping_records = get_customer_mappings_raw(6)
# delivery_location -> region (type 11)
dl_to_region = {}
# warehouse -> region (type 32)
wh_to_region = {}
for m in mapping_records:
    ot = str(m.get('order_type', '')).strip()
    dl = str(m.get('delivery_location', '')).strip() if m.get('delivery_location') else ''
    wh = str(m.get('warehouse', '')).strip() if m.get('warehouse') else ''
    region = str(m['region']).strip() if m['region'] else ''
    if ot == '11' and dl:
        dl_to_region[dl] = region
    elif ot == '32' and wh:
        wh_to_region[wh] = region

print(f'Delivery lookup: {len(dl_to_region)} entries')
print(f'Warehouse lookup: {len(wh_to_region)} entries')

# 3. 建立 ERP lookup: 送貨地點(AG) -> 訂單型態前綴
erp_location_to_type = {}
for _, row in erp_df.iterrows():
    ag = str(row['送貨地點']).strip() if pd.notna(row['送貨地點']) else ''
    am = str(row['訂單型態']).strip() if pd.notna(row['訂單型態']) else ''
    ot_prefix = am[:2] if len(am) >= 2 else ''
    if ag and ot_prefix:
        erp_location_to_type[ag] = ot_prefix

print(f'ERP location->type: {len(erp_location_to_type)} entries')
for loc, ot in sorted(erp_location_to_type.items()):
    print(f'  [{loc}] -> type {ot}')

# 4. Transit mapping
k_col = transit_df.columns[10]  # K column: "11訂單>送貨地點\n32訂單>倉庫"
d_col = transit_df.columns[3]   # D column: Location

print(f'\nK col name: [{k_col}]')
print(f'D col name: [{d_col}]')

def get_transit_region(row):
    location = str(row[d_col]).strip() if pd.notna(row[d_col]) else ''
    k_val = str(row[k_col]).strip() if pd.notna(row[k_col]) else ''

    # Step 1: Transit D -> ERP AG -> 訂單型態
    ot_prefix = erp_location_to_type.get(location, '')

    # Step 2: 根據訂單型態，K 值查 mapping
    if ot_prefix == '11':
        # K = 送貨地點 -> 查 delivery_location lookup
        return dl_to_region.get(k_val, '')
    elif ot_prefix == '32':
        # K = 倉庫 -> 查 warehouse lookup
        return wh_to_region.get(k_val, '')
    else:
        # fallback: 兩邊都試
        region = dl_to_region.get(k_val, '') or wh_to_region.get(k_val, '')
        return region

transit_df['客戶需求地區'] = transit_df.apply(get_transit_region, axis=1)

# 5. 統計
total = len(transit_df)
matched = (transit_df['客戶需求地區'] != '').sum()
print(f'\n=== Transit 匹配結果 ===')
print(f'總行數: {total}')
print(f'匹配成功: {matched} ({matched/total*100:.1f}%)')
print(f'未匹配: {total - matched}')

# 6. 顯示結果
print(f'\n=== 全部結果 ===')
for idx, row in transit_df.iterrows():
    loc = row[d_col]
    k = row[k_col]
    region = row['客戶需求地區']
    print(f'  D={loc} | K={k} | region={region}')
    if idx >= 14:
        print(f'  ... ({total - 15} more rows)')
        break

# 7. 輸出
output = os.path.join(os.path.dirname(__file__), 'liteon_integrated_transit.xlsx')
transit_df.to_excel(output, index=False)
print(f'\n結果已輸出到 test/liteon_integrated_transit.xlsx')
