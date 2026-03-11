# FORECAST 數據處理系統 — 光寶科技客製化擴展 PRD（工程師版）

**文件版本**: v1.0
**建立日期**: 2026-03-11
**專案名稱**: FORECAST 數據處理系統 — 光寶科技 (Liteon) 客製化擴展
**機密等級**: 內部工程文件
**客戶代碼**: user_id=6, username=liteon

---

## 1. 專案概述

### 1.1 專案背景

在現有 FORECAST 數據處理系統上新增第三個客戶支援（光寶科技 Liteon），前兩個為廣達 (Quanta, user_id=3) 與和碩 (Pegatron, user_id=5)。光寶的業務邏輯與前兩者有顯著差異，包含雙訂單類型、多檔案 Forecast、以及 ETD/ETA 日期算法選擇。

### 1.2 變更範圍總覽

| 新增/修改 | 檔案 | 行數範圍 | 說明 |
|:---------:|------|---------|------|
| 新增 | `liteon_forecast_processor.py` | 全檔 (561行) | 光寶專用 Forecast 處理器 |
| 修改 | `app.py` | ~2756-2837 | ERP Mapping 整合（Liteon 分支） |
| 修改 | `app.py` | ~2942-3006 | Transit Mapping 整合（Liteon 分支） |
| 修改 | `app.py` | ~3224-3326 | Forecast 處理路由（Liteon 分支） |
| 修改 | `static/js/mapping.js` | 多處 | Mapping 介面擴展四欄位 |
| 修改 | `templates/mapping.html` | thead 區域 | 動態表頭支援 |
| 修改 | `database.py` | `save_customer_mappings_list()` | 新增欄位存儲 |
| 新增 | `test/test_liteon_forecast.py` | 全檔 | 測試腳本 |
| 新增 | `test/test_liteon_mapping.py` | 全檔 | ERP Mapping 測試 |
| 新增 | `test/test_liteon_transit_mapping.py` | 全檔 | Transit Mapping 測試 |

### 1.3 客戶識別機制

系統在三個層級進行客戶識別：

```python
# app.py 內部識別邏輯
target_username = user['username'].lower()
is_pegatron = target_username == 'pegatron'
is_liteon = target_username == 'liteon'

# IT 測試模式
is_pegatron = processor_user_id == 5
is_liteon = processor_user_id == 6
```

---

## 2. 功能需求詳細

### 2.1 檔案上傳（階段一）

#### 2.1.1 ERP 淨需求

光寶 ERP 的關鍵欄位（依 Excel 欄位代號）：

| ERP 欄位 | 代號 | 用途 |
|----------|------|------|
| 客戶簡稱 | D | Mapping 比對主鍵之一 |
| 送貨地點 | AG | 訂單型態 11 的比對鍵 |
| 倉庫 | AL | 訂單型態 32 的比對鍵 |
| 訂單型態 | AM | 區分 "11一般訂單" / "32HUB補貨單" |
| 客戶料號 | — | 比對 Forecast 料號 (Material) |
| 排程出貨日期 | — | 日期計算起始點 |
| 淨需求 | — | 填入 Forecast 的數量來源（×1000） |

#### 2.1.2 Forecast 預測

- Sheet 名稱：`Daily+Weekly+Monthly`
- C1 = Plant code（廠區代碼，如 "15K0"）
- E1 = Buyer code（採購員代碼，如 "P43"）
- Row 7 = 日期標頭
  - K~AO (col 11~41)：Daily（每日，最多 31 天）
  - AP~BK (col 42~63)：Weekly（每週，最多 22 週）
  - BL~BQ (col 64~69)：Monthly（每月，最多 6 個月）
- Row 8+：每個料號 3 列一組
  - B 欄 = Material（料號）
  - C 欄 = Data Measures（"Demand" / "Commit" / "Accumulate Shortage"）
- 多檔上傳：每個 Plant+Buyer 組合各一份，測試數據共 23 個檔案

#### 2.1.3 Transit 在途

光寶 Transit 的關鍵欄位（依 Excel 欄位索引）：

| Transit 欄位 | 索引 | 用途 |
|-------------|------|------|
| Location | D (index 3) | 反查 ERP 送貨地點 → 得知訂單型態 |
| K 欄 | K (index 10) | 11 訂單=送貨地點 / 32 訂單=倉庫 |
| Ordered Item | — | 比對 Forecast 料號 (Material) |
| Qty | — | 填入 Forecast 的數量來源（×1000） |
| ETA | — | 填入日期依據 |

---

### 2.2 數據清理（階段二）

使用共用清理邏輯，清除 Forecast 的 Commit 列和 Accumulate Shortage 列中的舊數據。保留 Excel 格式（字型、框線、合併儲存格）。

---

### 2.3 Mapping 整合（階段三）

#### 2.3.1 Mapping 資料庫欄位

光寶 Mapping 比現有客戶多 4 個欄位：

| 欄位名 | DB Key | 類型 | 說明 |
|--------|--------|------|------|
| 客戶簡稱 | `customer_name` | string | 比對 ERP 客戶簡稱 (D欄) |
| 訂單型態 | `order_type` | "11" / "32" | 區分一般訂單與 HUB 調撥 |
| 送貨地點 | `delivery_location` | string | 訂單型態 11 的第二比對鍵 |
| 倉庫 | `warehouse` | string | 訂單型態 32 的第二比對鍵 |
| 廠區 | `region` | string | 對應 Forecast C1 (Plant code) |
| 排程斷點 | `schedule_breakpoint` | string | 週中斷點（如 "禮拜一"） |
| ETD | `etd` | string | ETD 文字（如 "下週四"） |
| ETA | `eta` | string | ETA 文字（如 "下下週二"） |
| 日期算法 | `date_calc_type` | "ETD" / "ETA" | 選擇使用哪個日期邏輯 |
| Transit 需求 | `requires_transit` | boolean | 是否需要 Transit 數據 |

#### 2.3.2 ERP Mapping 演算法

```
app.py:2756-2837

輸入：ERP DataFrame + Mapping Records
輸出：ERP DataFrame 新增 5 欄（客戶需求地區、排程出貨日期斷點、ETD、ETA、日期算法）

步驟：
1. 從 DB 載入 mapping_records（get_customer_mappings_raw(user_id)）
2. 建立兩個 lookup dict：
   - liteon_lookup_11: (customer_name, delivery_location) → mapping_values
   - liteon_lookup_32: (customer_name, warehouse) → mapping_values
3. 動態查找 ERP 欄位名稱（find_column_by_name）
4. 對每一筆 ERP 數據：
   a. 讀取訂單型態 (AM欄) → 取前 2 字元 ("11" or "32")
   b. if "11" → 用 (客戶簡稱, 送貨地點) 查 liteon_lookup_11
   c. if "32" → 用 (客戶簡稱, 倉庫) 查 liteon_lookup_32
   d. 匹配成功 → 填入 region, schedule_breakpoint, etd, eta, date_calc_type
```

#### 2.3.3 Transit Mapping 演算法

```
app.py:2942-3006

輸入：Transit DataFrame + ERP DataFrame + Mapping Records
輸出：Transit DataFrame 新增 客戶需求地區 欄位

步驟：
1. 從 Mapping 建立 dl_to_region (送貨地點→region) 和 wh_to_region (倉庫→region)
2. 從 ERP 建立 erp_location_to_type (ERP AG送貨地點 → 訂單型態前綴 11/32)
3. 對每一筆 Transit 數據：
   a. 讀取 Transit D (Location)
   b. 用 Location 查 erp_location_to_type → 得到訂單型態 (11/32)
   c. 讀取 Transit K 欄
   d. if "11" → 用 K 值查 dl_to_region → 得到 region
   e. if "32" → 用 K 值查 wh_to_region → 得到 region
   f. fallback: 兩邊都試
```

---

### 2.4 Forecast 處理（階段四）

#### 2.4.1 處理器入口

```
app.py:3226-3326

入口：/run_forecast route, is_liteon=True
流程：
  for each forecast_file in multi_cleaned_files:
    1. 讀取 C1 (Plant) + E1 (Buyer) → 產生 output_filename
    2. new LiteonForecastProcessor(forecast_file, erp_file, transit_file, ...)
    3. processor.process_all_blocks()
    4. 累計統計
    5. 分配狀態自動回寫（跨檔案共享 ERP/Transit）
```

#### 2.4.2 LiteonForecastProcessor 完整流程

```
liteon_forecast_processor.py

Class: LiteonForecastProcessor

__init__(forecast_file, erp_file, transit_file, output_folder, output_filename)
  - 初始化所有路徑和統計計數器
  - pending_changes: [] (延遲寫入佇列)

process_all_blocks():  [主入口]
  ① _load_forecast()
     - openpyxl.load_workbook(forecast_file)
     - 讀取 C1 → self.plant_code
     - _parse_date_headers() → self.date_map {col_index: date}
     - _build_material_index() → self.material_commit_rows {material: row}

  ② _load_data()
     - pd.read_excel(erp_file) → self.erp_df
     - pd.read_excel(transit_file) → self.transit_df
     - 初始化「已分配」欄位

  ③ _process_transit()  [先處理 Transit]
     - 比對: transit[客戶需求地區] == self.plant_code
     - 比對: transit[Ordered Item] == forecast[Material]
     - 解析 transit[ETA] → target_date
     - _find_date_column(target_date) → col
     - _add_change(commit_row, col, qty * 1000)
     - 標記已分配 ✓

  ④ _process_erp()  [再處理 ERP]
     - 比對: erp[客戶需求地區] == self.plant_code
     - 比對: erp[客戶料號] == forecast[Material]
     - _calculate_erp_target_date():
       a. 讀取 排程出貨日期 → schedule_date
       b. 讀取 排程出貨日期斷點 → breakpoint_text
       c. _get_week_end_by_breakpoint(schedule_date, breakpoint_text) → week_end
       d. 讀取 日期算法(ETD/ETA) → 選擇 ETD 或 ETA 文字
       e. _calculate_target_from_text(week_end, date_text) → target_date
       f. 防護: target_date >= schedule_date (否則 return None)
     - _find_date_column(target_date) → col
     - _add_change(commit_row, col, qty * 1000)
     - 標記已分配 ✓

  ⑤ _apply_changes()
     - 遍歷 pending_changes
     - 若原有值為數字 → 累加
     - 否則 → 直接寫入

  ⑥ _save_file()
     - wb.save(output_path)

  ⑦ _save_allocation_status()
     - erp_df.to_excel(erp_file) → 回寫已分配狀態
     - transit_df.to_excel(transit_file) → 回寫已分配狀態
```

#### 2.4.3 日期計算核心演算法

**_get_week_end_by_breakpoint(schedule_date, breakpoint_text)**

```python
# 將斷點文字轉為 weekday 數字
weekday_map = {'週一': 0, '禮拜一': 0, '星期一': 0, ...}
target_weekday = weekday_map[breakpoint_text]

# 從排程出貨日期往後（含當天）找到下一個斷點日
current_weekday = schedule_date.weekday()
days_ahead = (target_weekday - current_weekday) % 7
week_end = schedule_date + timedelta(days=days_ahead)

# 注意: days_ahead=0 表示排程出貨日當天就是斷點日
```

**_calculate_target_from_text(week_end, date_text)**

```python
# 解析文字
"本週X"   → weeks_offset = 0
"下週X"   → weeks_offset = 1
"下下週X" → weeks_offset = 2

# 目標 weekday
weekday_map = {'一': 0, '二': 1, '三': 2, '四': 3, '五': 4, '六': 5, '日': 6}
target_weekday = weekday_map[weekday_char]

# 核心公式（已修正）
breakpoint_weekday = week_end.weekday()
days_diff = (target_weekday - breakpoint_weekday) % 7
if days_diff > 0:
    days_diff -= 7  # 目標在斷點日之前（同一周期內）

return week_end + timedelta(days=7 * weeks_offset + days_diff)
```

**公式解釋**：
- 斷點日是一個週期的最後一天
- 目標 weekday 必定在斷點日或之前（同一週期內）
- `days_diff > 0` 代表原始差值為正（目標 weekday 在斷點之後），需減 7 拉回同一週期
- `days_diff == 0` 代表目標就是斷點日本身
- `days_diff < 0` 代表目標在斷點日之前，已是正確位置

**驗算範例**：
```
排程出貨日期: 3/10 (週二)
斷點: 禮拜一
→ week_end = 3/16 (週一)

ETA = "下週四"
→ weeks_offset=1, target_weekday=3(週四)
→ breakpoint_weekday=0(週一)
→ days_diff = (3-0)%7 = 3, 3>0 → 3-7 = -4
→ target = 3/16 + 7*1 + (-4) = 3/16 + 3 = 3/19 (週四) ✓

ETD = "本週五"
→ weeks_offset=0, target_weekday=4(週五)
→ days_diff = (4-0)%7 = 4, 4>0 → 4-7 = -3
→ target = 3/16 + 0 + (-3) = 3/13 (週五) ✓
  3/13 >= 3/10 (排程出貨日期) ✓

驗算2：排程出貨日期: 3/17 (週二), 斷點: 禮拜四
→ week_end = 3/19 (週四)
ETA = "下週二"
→ weeks_offset=1, target_weekday=1(週二)
→ breakpoint_weekday=3(週四)
→ days_diff = (1-3)%7 = 5, 5>0 → 5-7 = -2
→ target = 3/19 + 7 + (-2) = 3/24 (週二) ✓
```

#### 2.4.4 日期對應策略 (_find_date_column)

```python
# Step 1: Daily 精確匹配 (col 11~41)
for col in daily_range:
    if date_map[col] == target_date: return col

# Step 2: Weekly 範圍匹配 (col 42~63)
for col in weekly_range:
    week_start = date_map[col]
    week_end = week_start + 6 days
    if week_start <= target_date <= week_end: return col

# Step 3: Monthly 月份匹配 (col 64~69)
for col in monthly_range:
    if date_map[col].year == target.year and date_map[col].month == target.month: return col

# Step 4: return None（無匹配，跳過此筆）
```

#### 2.4.5 分配追蹤機制

```
多檔案處理的分配追蹤靠 ERP/Transit DataFrame 的「已分配」欄位：

Forecast #1 處理完 → erp_df 部分行 已分配='✓'
  ↓ _save_allocation_status() → erp_df.to_excel(erp_file)

Forecast #2 開始 → _load_data() → pd.read_excel(erp_file)
  ↓ 讀到已分配='✓' 的行 → 跳過

每個 Forecast 檔案處理完都會回寫 ERP/Transit 的 Excel，
下一個檔案重新讀取時就能看到前面檔案的分配結果。
```

#### 2.4.6 數量填入規則

| 規則 | 實現 |
|------|------|
| ×1000 單位轉換 | `fill_value = qty * 1000` |
| 同儲存格累加 | `_add_change()` 內部累加 |
| 保留原值 | `_apply_changes()`: `cell.value = current + change['value']` |
| qty <= 0 跳過 | ERP 處理中 `if qty <= 0: continue` |

---

### 2.5 結果下載（階段五）

#### 2.5.1 輸出檔案命名

```python
# app.py:3245-3259
_tmp_wb = openpyxl.load_workbook(forecast_file, read_only=True)
_tmp_ws = _tmp_wb['Daily+Weekly+Monthly']
plant_code = str(_tmp_ws.cell(row=1, column=3).value or '').strip()  # C1
buyer_code = str(_tmp_ws.cell(row=1, column=5).value or '').strip()  # E1

if plant_code and buyer_code:
    output_filename = f'forecast_{plant_code}_{buyer_code}.xlsx'
elif plant_code:
    output_filename = f'forecast_{plant_code}.xlsx'
else:
    output_filename = f'forecast_result_{file_num}.xlsx'
```

#### 2.5.2 API 回傳格式

```json
{
  "success": true,
  "message": "FORECAST處理完成：23 個檔案",
  "multi_file": true,
  "files": [
    {
      "input": "cleaned_forecast_1.xlsx",
      "output": "forecast_15K0_P43.xlsx",
      "erp_filled": 52,
      "transit_filled": 3,
      "file_size": 245760
    }
  ],
  "file_count": 23,
  "success_count": 23,
  "total_erp_filled": 1240,
  "total_erp_skipped": 19,
  "total_transit_filled": 43,
  "total_transit_skipped": 1,
  "transit_file_skipped": false
}
```

前端 `main.js:updateDownloadSectionMultiFile(files)` 處理多檔案下載 UI。

---

## 3. Mapping 介面設計

### 3.1 前端擴展

**mapping.js 新增函式**：

| 函式 | 說明 |
|------|------|
| `hasLiteonFields()` | 偵測 mapping data 中是否有 order_type/delivery_location/warehouse 欄位 |
| 動態表頭渲染 | `renderMappingTableList()` 中根據 `hasLiteonFields()` 展開 4 個額外欄位 |
| 儲存時攜帶新欄位 | `saveCurrentPageEdits()` 抓取 order_type, delivery_location, warehouse, date_calc_type |

**mapping.html 修改**：
- `<tr id="mapping-thead-row">` 動態表頭支援

### 3.2 CSS 擴展

```css
.liteon-expanded {
    /* 表格寬度自動擴展以容納額外 4 欄 */
}
```

### 3.3 資料庫擴展

`database.py:save_customer_mappings_list()` 新增存儲欄位：
- `order_type`
- `delivery_location`
- `warehouse`
- `date_calc_type`

---

## 4. 測試數據與結果

### 4.1 測試數據位置

```
processed/6/20260311_203430/
├── integrated_erp.xlsx          # 已整合 Mapping 的 ERP
├── integrated_transit.xlsx       # 已整合 Mapping 的 Transit
├── cleaned_forecast_1.xlsx       # 清理後 Forecast #1
├── cleaned_forecast_2.xlsx       # 清理後 Forecast #2
├── ...
└── cleaned_forecast_23.xlsx      # 清理後 Forecast #23
```

### 4.2 測試結果

| 指標 | 數值 |
|------|------|
| Forecast 檔案數 | 23 |
| 不重複 Plant 數 | 19 |
| Transit 填入 | 43/44（97.7%） |
| Transit 未匹配 | 1 筆（料號 8A0KA5 不在任何 Forecast） |
| ERP 填入 | 1,240+ 筆 |
| 處理時間 | < 30 秒（23 檔案） |

### 4.3 測試腳本

```
test/test_liteon_forecast.py     # 完整 Forecast 處理測試
test/test_liteon_mapping.py      # ERP Mapping 邏輯測試
test/test_liteon_transit_mapping.py  # Transit Mapping 邏輯測試
```

---

## 5. 已知限制與注意事項

### 5.1 日期計算限制

| 限制 | 說明 |
|------|------|
| 文字格式 | 僅支援 "本週X"、"下週X"、"下下週X"、"下下下週X" 格式 |
| 同義詞 | 支援 "週/禮拜/星期" 三種寫法 |
| 日期防護 | 目標日期 < 排程出貨日期 → 跳過（不填入） |

### 5.2 分配追蹤限制

| 限制 | 說明 |
|------|------|
| 檔案回寫 | 每個 Forecast 處理完都回寫 ERP/Transit Excel，重新讀取 |
| 格式損失 | ERP/Transit 回寫時使用 pandas to_excel，會失去原始格式 |
| 並行不安全 | 多用戶不會同時處理同一批數據（各有獨立目錄） |

### 5.3 欄位名稱依賴

| 元件 | 依賴的欄位名稱 |
|------|---------------|
| ERP Mapping | 客戶簡稱、送貨地點、訂單型態、倉庫（由 find_column_by_name 動態查找） |
| Transit Mapping | 固定用 index 3 和 index 10 |
| Forecast Processor | 客戶需求地區、客戶料號、排程出貨日期、排程出貨日期斷點、ETD、ETA、日期算法、淨需求、已分配 |
| Forecast Sheet | C1=Plant, Row 7=dates, Column B=Material, Column C=Measures |

---

## 6. 開發歷程中的重要修正

### 6.1 日期計算公式修正

**原始公式**（有 bug）：
```python
# 從 week_end 往後找目標 weekday → 多算一週
days_diff = (target_weekday - base_weekday) % 7
return week_end + timedelta(days=7 * weeks_offset + days_diff)
```

**修正後公式**：
```python
# 目標在斷點日「之前或當天」（同一週期內）
days_diff = (target_weekday - breakpoint_weekday) % 7
if days_diff > 0:
    days_diff -= 7
return week_end + timedelta(days=7 * weeks_offset + days_diff)
```

**根因**：斷點日是週期的最後一天，目標 weekday 在同一週期內，不應往後推。

### 6.2 ERP 數量欄位修正

- 原始：使用 `未交量` 欄位
- 修正：改為 `淨需求` 欄位

### 6.3 ERP 錯誤處理強化

- 原始：單一 ERP 行出錯 → 整個檔案失敗（連同 Transit 填入都丟失）
- 修正：per-row try/except，個別行錯誤不影響其他行

### 6.4 已分配欄位型態修正

- 問題：`FutureWarning` — 在 float64 column 設定 string '✓'
- 修正：`.astype(str).replace('nan', '')` 統一為字串型

### 6.5 日期安全防護

- 新增：`if target_date < schedule_date: return None`
- 原因：計算結果可能因為本週/下週的邊界條件導致日期早於排程出貨日期

---

## 7. 術語與常數參照

### 7.1 Forecast Sheet 常數

```python
SHEET_NAME = 'Daily+Weekly+Monthly'
PLANT_CELL_ROW = 1, PLANT_CELL_COL = 3      # C1
DATE_HEADER_ROW = 7
DATA_START_ROW = 8
MATERIAL_COL = 2                              # Column B
DATA_MEASURES_COL = 3                         # Column C
DAILY_START_COL = 11  (K)  ~ DAILY_END_COL = 41  (AO)
WEEKLY_START_COL = 42 (AP) ~ WEEKLY_END_COL = 63  (BK)
MONTHLY_START_COL = 64 (BL) ~ MONTHLY_END_COL = 69 (BQ)
```

### 7.2 Weekday 對照

```python
weekday_map = {
    '週一/禮拜一/星期一': 0,  # Monday
    '週二/禮拜二/星期二': 1,  # Tuesday
    '週三/禮拜三/星期三': 2,  # Wednesday
    '週四/禮拜四/星期四': 3,  # Thursday
    '週五/禮拜五/星期五': 4,  # Friday
    '週六/禮拜六/星期六': 5,  # Saturday
    '週日/禮拜日/星期日/禮拜天/週天': 6,  # Sunday
}
```
