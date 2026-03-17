# Liteon Forecast 日期範圍差異分析報告

**發現日期**: 2026/03/13
**影響範圍**: Liteon Forecast 合併處理模式
**嚴重程度**: 中（數據正確性影響）
**狀態**: 已修復

---

## 1. 問題摘要

在實作 Liteon Forecast「先合併再處理」功能時，發現合併模式的 ERP 填入數量（1270 筆）與逐檔模式（1240 筆）不一致，差異 30 筆。經調查確認為**不同 Forecast 檔案的 Daily 起始日期不同**所導致的邏輯問題。

---

## 2. 發現過程

| 處理模式 | ERP 填入 | Transit 填入 |
|---------|---------|-------------|
| 逐檔模式（23 個檔案分別處理） | 1,240 筆 | 43 筆 |
| 合併模式（23 個檔案合併後處理） | 1,270 筆 | 43 筆 |
| **差異** | **+30 筆** | 0 |

逐步排除後確認：
- 料號匹配結果完全一致（兩種模式皆 1,493 筆匹配）
- ERP 資料經 pandas 讀寫後無型態異常
- **根本原因為日期結構差異**

---

## 3. 根本原因

### 3.1 Forecast 檔案的日期結構

每個 Liteon Forecast 檔案的 `Daily+Weekly+Monthly` sheet 有三個日期區段：

```
| Daily (BY天) | GAP | Weekly (BY周) | Monthly (BY月) |
| K欄 ~ AO欄   |     | AP欄 ~ BK欄   | BL欄 ~ BQ欄   |
```

- **Daily**: 31 個日欄位（逐日）
- **Weekly**: 週彙總欄位（每週一個）
- **Monthly**: 月彙總欄位
- **GAP**: Daily 結束日到 Weekly 開始日之間的空白期，此區間的 ERP 資料**無法被填入任何欄位**

### 3.2 不同檔案的起始日不同

調查 23 個 Forecast 檔案發現，Daily 起始日並非統一：

| 檔案 | Plant | Daily 起始日 | Daily 結束日 | Weekly 起始日 | GAP 區間 | GAP 天數 |
|------|-------|------------|------------|-------------|---------|---------|
| File 1 | 15K0 | 2026/03/09 | 2026/04/08 | 2026/04/13 | 04/09 ~ 04/12 | **4 天** |
| File 6 | 42P0 | 2026/03/07 | 2026/04/06 | 2026/04/13 | 04/07 ~ 04/12 | **6 天** |
| File 19 | 429E | 2026/03/07 | 2026/04/06 | 2026/04/13 | 04/07 ~ 04/12 | **6 天** |

> 所有檔案的 **Weekly 起始日都是 2026/04/13（週一）**，但 Daily 結束日因起始日不同而不同。

### 3.3 合併後的日期結構問題

合併函式以第一個檔案（File 1, Plant 15K0）的 Row 7 日期作為統一標頭：

```
合併檔案 Daily 範圍: 2026/03/09 ~ 2026/04/08 (來自 File 1)
```

這導致 March 7 起始的 Plant（42P0, 42V0, 429E, 429H, 429L）在合併檔案中「多了」April 7-8 兩天的 Daily 欄位：

```
Plant 42P0 原始: Daily 到 04/06 → GAP 04/07~04/12 (6天)
Plant 42P0 合併: Daily 到 04/08 → GAP 04/09~04/12 (4天)  ← 多了 04/07, 04/08
```

### 3.4 影響

ERP 中目標日期為 2026/04/07 和 2026/04/08、屬於 March-7-start Plant 的條目：

- **逐檔模式**: 目標日期落入 GAP → 找不到欄位 → **不填入**（正確行為）
- **合併模式**: 目標日期在 Daily 範圍內 → 找到欄位 → **被填入**（錯誤行為）

這就是多出的 30 筆 ERP 填入的來源。

---

## 4. 修復方案

在合併過程中記錄每個 Plant 的原始 Daily 結束日期，並在填入時限制搜尋範圍。

### 4.1 合併函式 (`merge_liteon_forecast_files`)

新增 `plant_daily_end_dates` 字典，記錄每個 Plant 的原始 Daily 結束日期：

```python
plant_daily_end_dates = {
    '15K0': date(2026, 4, 8),   # March 9 起始 → Daily 到 April 8
    '42P0': date(2026, 4, 6),   # March 7 起始 → Daily 到 April 6
    '42V0': date(2026, 4, 6),
    '429E': date(2026, 4, 6),
    # ... 其他 Plant
}
```

### 4.2 Processor 日期搜尋 (`_find_date_column`)

新增 `plant` 參數，合併模式下限制 Daily 欄位搜尋：

```python
def _find_date_column(self, target_date, plant=None):
    # 取得此 Plant 的原始 Daily 結束日期
    plant_daily_limit = self.plant_daily_end_dates.get(plant)

    # Step 1: Daily — 跳過超過此 Plant 原始範圍的日期
    for col, date_obj in self.date_map.items():
        if plant_daily_limit and date_obj > plant_daily_limit:
            continue  # 超過此 Plant 的 Daily 範圍，跳過
        if date_obj == target_date:
            return col

    # Step 2: Weekly fallback
    # Step 3: Monthly fallback
```

### 4.3 修復結果

| 處理模式 | ERP 填入 | Transit 填入 |
|---------|---------|-------------|
| 逐檔模式 | 1,240 筆 | 43 筆 |
| 合併模式（修復後） | **1,240 筆** | 43 筆 |
| 差異 | **0** | 0 |

---

## 5. 結論與建議

### 5.1 結論

此問題為**邏輯流程問題**，非算法錯誤。Liteon 不同 Plant 的 Forecast 檔案因業務需求可能有不同的 Daily 起始日期，合併時必須保留每個 Plant 的原始日期範圍資訊，避免因統一標頭而改變了 GAP 的覆蓋範圍。

### 5.2 建議

1. **持續監控**: 未來若 Liteon 的 Forecast 格式變動（例如 Daily 欄位數量改變），需同步更新合併邏輯
2. **確認 GAP 行為**: 建議與光寶確認：Daily-Weekly 之間的 GAP 區間，ERP 資料是否應該 fall through 到 Weekly 彙總欄位，而非直接跳過。目前邏輯為「找不到 Daily → 找 Weekly → 找 Monthly」，GAP 內的日期會 fall through 到 Weekly
3. **Weekly 起始日一致性**: 目前所有檔案的 Weekly 起始日都是同一天（2026/04/13），如果未來不同檔案的 Weekly 起始日也不同，可能需要類似的處理機制

---

## 6. 影響檔案

| 檔案 | 修改內容 |
|------|---------|
| `app.py` — `merge_liteon_forecast_files()` | 收集 `plant_daily_end_dates`，回傳給呼叫端 |
| `app.py` — `run_forecast()` | 將 `plant_daily_end_dates` 傳遞給 Processor |
| `liteon_forecast_processor.py` — `__init__()` | 接收 `plant_daily_end_dates` 參數 |
| `liteon_forecast_processor.py` — `_find_date_column()` | 新增 `plant` 參數，限制 Daily 搜尋範圍 |
| `liteon_forecast_processor.py` — `_process_erp()` | 傳入 `plant=region` |
| `liteon_forecast_processor.py` — `_process_transit()` | 傳入 `plant=region` |
