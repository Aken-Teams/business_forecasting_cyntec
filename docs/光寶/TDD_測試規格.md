# 光寶 Forecast 系統 - TDD 測試規格

> Test-Driven Development 測試文件
> 以單元測試為核心，覆蓋每個函數的輸入/輸出/邊界條件

---

## 目錄

1. [日期解析模組](#1-日期解析模組)
2. [斷點與錨點計算](#2-斷點與錨點計算)
3. [目標日期計算（ETD/ETA 文字解析）](#3-目標日期計算etdeta-文字解析)
4. [日期欄位查找](#4-日期欄位查找)
5. [ERP 目標日期整合計算](#5-erp-目標日期整合計算)
6. [數量處理與累加](#6-數量處理與累加)
7. [料號索引建立](#7-料號索引建立)
8. [客戶映射前端解析](#8-客戶映射前端解析)
9. [ERP/Transit 映射查表](#9-erptransit-映射查表)
10. [Forecast 清理與合併](#10-forecast-清理與合併)

---

## 1. 日期解析模組

**目標函數**: `_parse_date_value(val)`

### Test Suite: `TestParseDateValue`

```
TC-1.1  輸入 None → 回傳 None
TC-1.2  輸入 float NaN → 回傳 None
TC-1.3  輸入 datetime(2026,3,30) → 回傳 date(2026,3,30)
TC-1.4  輸入 date(2026,3,30) → 回傳 date(2026,3,30)
TC-1.5  輸入 "2026/03/30" → 回傳 date(2026,3,30)
TC-1.6  輸入 "2026-03-30" → 回傳 date(2026,3,30)
TC-1.7  輸入 "03/30/2026" → 回傳 date(2026,3,30)
TC-1.8  輸入 "" → 回傳 None
TC-1.9  輸入 "nan" → 回傳 None
TC-1.10 輸入 "nat" → 回傳 None
TC-1.11 輸入 "abc" → 回傳 None
TC-1.12 輸入 pd.Timestamp("2026-03-30") → 回傳 date(2026,3,30)
```

---

## 2. 斷點與錨點計算

**目標函數**: `_get_week_end_by_breakpoint(schedule_date, breakpoint_text)`

### Test Suite: `TestGetWeekEndByBreakpoint`

**基本功能**:
```
TC-2.1  schedule=3/10(二), breakpoint="禮拜一" → 3/16(一)
        # 向前找最近的週一 = 3/16
TC-2.2  schedule=3/10(二), breakpoint="禮拜五" → 3/13(五)
        # 向前找最近的週五 = 3/13
TC-2.3  schedule=3/10(二), breakpoint="禮拜二" → 3/10(二)
        # 剛好是當天 = 3/10
```

**on_breakpoint 邊界條件**（排程日 == 斷點日）:
```
TC-2.4  schedule=3/16(一), breakpoint="禮拜一" → 3/16(一)
        # 排程剛好在斷點日上 → week_end == schedule_date
TC-2.5  schedule=3/13(五), breakpoint="禮拜五" → 3/13(五)
        # 排程剛好在斷點日上
TC-2.6  schedule=3/30(一), breakpoint="禮拜一" → 3/30(一)
        # 原始 bug 案例: schedule == week_end
```

**不同斷點 weekday**:
```
TC-2.7  schedule=3/11(三), breakpoint="禮拜三" → 3/11(三)
TC-2.8  schedule=3/11(三), breakpoint="禮拜四" → 3/12(四)
TC-2.9  schedule=3/11(三), breakpoint="禮拜一" → 3/16(一)
```

**異常輸入**:
```
TC-2.10 breakpoint="" → 回傳 None
TC-2.11 breakpoint="無效值" → 回傳 None
TC-2.12 schedule=None → 回傳 None（或拋例外）
```

**中文變體**:
```
TC-2.13 breakpoint="週一" → 與 "禮拜一" 結果相同
TC-2.14 breakpoint="星期一" → 與 "禮拜一" 結果相同
```

---

## 3. 目標日期計算（ETD/ETA 文字解析）

**目標函數**: `_calculate_target_from_text(week_end, date_text, on_breakpoint=False)`

### Test Suite: `TestCalculateTargetFromText`

**正常情況（on_breakpoint=False，斷點是當週結束日）**:
```
TC-3.1  week_end=3/16(一), "本週五", on_bp=False
        → days_diff=(4-0)%7=4 → 4-7=-3 → 3/16-3 = 3/13(五)
TC-3.2  week_end=3/16(一), "下週五", on_bp=False
        → 3/16 + 7 - 3 = 3/20(五)
TC-3.3  week_end=3/16(一), "下週四", on_bp=False
        → days_diff=(3-0)%7=3 → 3-7=-4 → 3/16 + 7 - 4 = 3/19(四)
TC-3.4  week_end=3/16(一), "下下週二", on_bp=False
        → days_diff=(1-0)%7=1 → 1-7=-6 → 3/16 + 14 - 6 = 3/24(二)
TC-3.5  week_end=3/16(一), "下下下週三", on_bp=False
        → days_diff=(2-0)%7=2 → 2-7=-5 → 3/16 + 21 - 5 = 4/1(三)
```

**on_breakpoint=True（排程日==斷點日，斷點是新一週起始日）**:
```
TC-3.6  week_end=3/16(一), "下週四", on_bp=True
        → days_diff=(3-0)%7=3 → 保持 3 → 3/16 + 7 + 3 = 3/26(四)
TC-3.7  week_end=3/30(一), "下週四", on_bp=True
        → 3/30 + 7 + 3 = 4/9(四) ← 原始 bug 驗證案例
TC-3.8  week_end=3/16(一), "本週五", on_bp=True
        → days_diff=4 → 3/16 + 0 + 4 = 3/20(五)
TC-3.9  week_end=3/13(五), "下下下週三", on_bp=True
        → days_diff=(2-4)%7=5 → 3/13 + 21 + 5 = 4/8(三)
```

**days_diff=0 情況（目標 weekday == 斷點 weekday）**:
```
TC-3.10 week_end=3/16(一), "下週一", on_bp=False
        → days_diff=0 → 3/16 + 7 + 0 = 3/23(一)
TC-3.11 week_end=3/16(一), "下週一", on_bp=True
        → days_diff=0 → 3/16 + 7 + 0 = 3/23(一)
        # days_diff=0 時 on_breakpoint 不影響結果
```

**禮拜五斷點組合（429E/429H/429L/42P0/42V0）**:
```
TC-3.12 week_end=3/13(五), "下下週二", on_bp=False
        → days_diff=(1-4)%7=4 → 4-7=-3 → 3/13 + 14 - 3 = 3/24(二)
TC-3.13 week_end=3/13(五), "下下下週三", on_bp=False
        → days_diff=(2-4)%7=5 → 5-7=-2 → 3/13 + 21 - 2 = 4/1(三)
TC-3.14 week_end=3/13(五), "下下下週三", on_bp=True
        → days_diff=5 → 保持 5 → 3/13 + 21 + 5 = 4/8(三)
```

**直接日期字串**:
```
TC-3.15 date_text="2026/04/01" → 回傳 date(2026,4,1)
TC-3.16 date_text="無效文字" → 回傳 None
```

**中文變體**:
```
TC-3.17 "下禮拜四" → 與 "下週四" 結果相同
TC-3.18 "本禮拜五" → 與 "本週五" 結果相同
TC-3.19 "這週三" → 與 "本週三" 結果相同
```

---

## 4. 日期欄位查找

**目標函數**: `_find_date_column(target_date, plant=None)`

### Test Suite: `TestFindDateColumn`

**前置條件**: 建立 date_map 模擬（Daily: 3/9~4/8, Weekly: 4/13~8/31, Monthly: 9~11）

**Daily 精確匹配**:
```
TC-4.1  target=3/15 → 回傳對應 Daily col
TC-4.2  target=4/8  → 回傳最後一個 Daily col
TC-4.3  target=3/9  → 回傳第一個 Daily col
```

**Weekly 範圍匹配**:
```
TC-4.4  target=4/13 → 回傳第一個 Weekly col (4/13 週起)
TC-4.5  target=4/15 → 回傳 4/13 的 Weekly col (4/13~4/19)
TC-4.6  target=4/19 → 回傳 4/13 的 Weekly col (週末邊界)
TC-4.7  target=4/20 → 回傳 4/20 的 Weekly col (下一週)
```

**GAP Fallback（Daily 結束 ~ Weekly 開始之間的空白區）**:
```
TC-4.8  Daily end=4/8, Weekly start=4/13
        target=4/9  → 回傳第一個 Weekly col (GAP fallback)
TC-4.9  target=4/10 → 回傳第一個 Weekly col (GAP fallback)
TC-4.10 target=4/12 → 回傳第一個 Weekly col (GAP fallback)
```

**03/07-start 組的 GAP（429E/429H/429L/42P0/42V0）**:
```
TC-4.11 Daily end=4/6, Weekly start=4/13
        target=4/7  → 回傳第一個 Weekly col (6天 GAP)
TC-4.12 target=4/8  → 回傳第一個 Weekly col (6天 GAP)
```

**Monthly 匹配**:
```
TC-4.13 target=9/15 → 回傳 September Monthly col
TC-4.14 target=10/1 → 回傳 October Monthly col
```

**NOT_FOUND**:
```
TC-4.15 target=12/1 → 回傳 None (超出 Monthly 範圍)
TC-4.16 target=2/1  → 回傳 None (早於 Daily 開始)
```

**合併模式 Plant Daily Limit**:
```
TC-4.17 merged_mode=True, plant="429E", plant_daily_limit=4/6
        target=4/7 → 不匹配 Daily (超過 limit) → Weekly/GAP fallback
TC-4.18 merged_mode=True, plant="15K0", plant_daily_limit=4/8
        target=4/8 → 匹配 Daily (在 limit 內)
```

**None 輸入**:
```
TC-4.19 target=None → 回傳 None
```

---

## 5. ERP 目標日期整合計算

**目標函數**: `_calculate_erp_target_date(row, ...)`

### Test Suite: `TestCalculateERPTargetDate`

**完整流程測試（端到端單元）**:
```
TC-5.1  排程=3/10(二), 斷點=禮拜一, 日期算法=ETA, ETA=下週四
        → week_end=3/16, on_bp=False → 3/19(四)

TC-5.2  排程=3/30(一), 斷點=禮拜一, 日期算法=ETA, ETA=下週四
        → week_end=3/30, on_bp=True → 4/9(四)  ← 關鍵修復案例

TC-5.3  排程=3/16(一), 斷點=禮拜一, 日期算法=ETD, ETD=本週五
        → week_end=3/16, on_bp=True → 3/20(五)

TC-5.4  排程=3/10(二), 斷點=禮拜五, 日期算法=ETA, ETA=下下下週三
        → week_end=3/13, on_bp=False → 4/1(三)

TC-5.5  排程=3/13(五), 斷點=禮拜五, 日期算法=ETA, ETA=下下下週三
        → week_end=3/13, on_bp=True → 4/8(三)
```

**日期算法選擇**:
```
TC-5.6  日期算法="ETD" → 使用 ETD 欄位的文字
TC-5.7  日期算法="ETA" → 使用 ETA 欄位的文字
TC-5.8  日期算法="" → 預設嘗試 ETD，若空則 ETA
```

**安全防護（目標 < 排程）**:
```
TC-5.9  排程=4/1, 計算結果=3/28 → 回傳 None（目標早於排程）
```

**缺值處理**:
```
TC-5.10 排程出貨日期=NaN → 回傳 None
TC-5.11 斷點="" → 回傳 None
TC-5.12 ETD="" 且 ETA="" → 回傳 None
```

---

## 6. 數量處理與累加

**目標函數**: `_add_change(row, col, value)` 與 `_apply_changes()`

### Test Suite: `TestQuantityProcessing`

**基本填入**:
```
TC-6.1  add_change(8, 15, 5000) → pending_changes = [{row:8, col:15, value:5000}]
TC-6.2  qty=5 → fill_value = 5 * 1000 = 5000
```

**同 cell 累加**:
```
TC-6.3  add_change(8, 15, 3000) → add_change(8, 15, 2000)
        → pending_changes = [{row:8, col:15, value:5000}]  # 累加
TC-6.4  add_change(8, 15, 1000) → add_change(8, 16, 2000)
        → 兩筆不同 cell 分開存
```

**apply 到既有值**:
```
TC-6.5  cell 原值=10000, apply 5000 → cell = 15000（相加）
TC-6.6  cell 原值=None, apply 5000 → cell = 5000（新值）
TC-6.7  cell 原值="" (非數字), apply 5000 → cell = 5000（覆蓋）
```

**ERP 數量驗證**:
```
TC-6.8  qty=0 → 跳過不填入
TC-6.9  qty=-5 → 跳過不填入
TC-6.10 qty=0.5 → fill_value = 500
```

---

## 7. 料號索引建立

**目標函數**: `_build_material_index()`

### Test Suite: `TestBuildMaterialIndex`

**單檔模式**:
```
TC-7.1  Commit row for "768290" at row 9
        → material_commit_rows["768290"] = 9
TC-7.2  Demand row → 跳過不索引
TC-7.3  Accumulate Shortage row → 跳過不索引
TC-7.4  material="" → 跳過不索引
```

**合併模式**:
```
TC-7.5  Plant="2680", material="768290", Commit at row 5
        → material_commit_rows[("2680", "768290")] = 5
TC-7.6  不同 Plant 同料號 → 兩筆獨立索引
TC-7.7  Plant="" → 跳過不索引
```

---

## 8. 客戶映射前端解析

**目標函數**: `parseWeekDay(value)` (JavaScript)

### Test Suite: `TestParseWeekDay`

```
TC-8.1  "下週四" → {week: "下週", day: "四"}
TC-8.2  "本週五" → {week: "本週", day: "五"}
TC-8.3  "下下週二" → {week: "下下週", day: "二"}
TC-8.4  "下下下週三" → {week: "下下下週", day: "三"}  ← 修復後的案例
TC-8.5  "上週一" → {week: "上週", day: "一"}
TC-8.6  "" → {week: "", day: ""}
TC-8.7  null → {week: "", day: ""}
TC-8.8  "無效文字" → {week: "", day: ""}
```

**目標函數**: `combineWeekDay(week, day)` (JavaScript)

```
TC-8.9  ("下週", "四") → "下週四"
TC-8.10 ("", "四") → ""（week 為空回傳空）
TC-8.11 ("下週", "") → ""（day 為空回傳空）
TC-8.12 ("下下下週", "三") → "下下下週三"
```

**分頁保存**: `saveCurrentPageEdits()` (JavaScript)

```
TC-8.13 修改 Page 1 的 ETA → 切到 Page 2 → mappingList[page1_index].eta 已更新
TC-8.14 修改 Page 2 的 ETD → 保存 → 送出的 mapping_list 包含所有頁面資料
```

---

## 9. ERP/Transit 映射查表

**目標函數**: `build_liteon_lookup_tables()` (app.py)

### Test Suite: `TestBuildLiteonLookup`

**Type 11 查表**:
```
TC-9.1  mapping: customer="光-常州", order_type="11", delivery_location="常州", region="2680"
        → liteon_lookup_11[("光-常州", "常州")] = {region: "2680", ...}

TC-9.2  ERP row: 客戶簡稱="光-常州", 訂單型態="11一般訂單", 送貨地點="常州"
        → 匹配 lookup_11 → 填入 客戶需求地區="2680"
```

**Type 32 查表**:
```
TC-9.3  mapping: customer="光-429E)", order_type="32", warehouse="倉庫", region="429E"
        → liteon_lookup_32[("光-429E)", "倉庫")] = {region: "429E", ...}

TC-9.4  ERP row: 客戶簡稱="光-429E)", 訂單型態="32HUB補貨單", 倉庫="倉庫"
        → 匹配 lookup_32 → 填入 客戶需求地區="429E"
```

**訂單型態前綴解析**:
```
TC-9.5  "11一般訂單" → prefix="11"
TC-9.6  "32HUB補貨單" → prefix="32"
TC-9.7  "99其他" → 無匹配
```

**Transit 映射**:
```
TC-9.8  Transit row: 訂單型態="11", 送貨地點="常州"
        → dl_to_region["常州"] = "2680" → 客戶需求地區="2680"
TC-9.9  Transit row: 訂單型態="32", 倉庫="倉庫"
        → wh_to_region["倉庫"] = "429E" → 客戶需求地區="429E"
```

---

## 10. Forecast 清理與合併

### Test Suite: `TestForecastCleanup`

**Commit row 清零**:
```
TC-10.1 Commit row, cols J~BY 有值 → 清零後全為 0/None
TC-10.2 Demand row → 保持原值不動
TC-10.3 Accumulate Shortage row → 保持原值不動
```

### Test Suite: `TestForecastMerge`

**合併邏輯**:
```
TC-10.4 2 個檔案合併 → Row 1 = headers, Row 2+ = 所有資料
TC-10.5 合併後 Column A = Plant, Column B = Buyer Code
TC-10.6 所有原始欄位右移 2 位
TC-10.7 日期欄位對齊：不同檔案相同日期 → 合併到同一欄
TC-10.8 plant_daily_end_dates 正確記錄各 Plant 的 Daily 結束日期
```

**單檔 fallback**:
```
TC-10.9  只有 1 個檔案 → 不合併，直接處理
TC-10.10 cleaned 資料夾只有 1 個清理後檔案 → 單檔模式
```

---

## 測試資料設計

### 最小 ERP 測試集

| 排程出貨日期 | 斷點 | ETD | ETA | 日期算法 | 客戶需求地區 | 客戶料號 | 淨需求 | 預期結果 |
|---|---|---|---|---|---|---|---|---|
| 3/10(二) | 禮拜一 | 本週五 | 下週四 | ETA | 2680 | MAT001 | 5 | 3/19(四) |
| 3/16(一) | 禮拜一 | 本週五 | 下週四 | ETA | 2680 | MAT001 | 10 | 3/26(四) |
| 3/30(一) | 禮拜一 | 本週五 | 下週四 | ETA | 2680 | MAT001 | 15 | **4/9(四)** |
| 3/10(二) | 禮拜五 | 下下週二 | 下下下週三 | ETA | 429E | MAT002 | 8 | 4/1(三) |
| 3/13(五) | 禮拜五 | 下下週二 | 下下下週三 | ETA | 429E | MAT002 | 12 | **4/8(三)** |
| 3/16(一) | 禮拜一 | 下週一 | 下週四 | ETD | 15K1 | MAT003 | 20 | 3/23(一) |

> **粗體** = on_breakpoint 案例

---

## 測試覆蓋率目標

| 模組 | 函數數 | 測試案例數 | 目標覆蓋率 |
|---|---|---|---|
| 日期解析 | 1 | 12 | 100% |
| 斷點計算 | 1 | 14 | 100% |
| 目標日期計算 | 1 | 19 | 100% |
| 日期欄位查找 | 1 | 19 | 100% |
| ERP 整合計算 | 1 | 12 | 100% |
| 數量處理 | 2 | 10 | 100% |
| 料號索引 | 1 | 7 | 100% |
| 前端解析 | 3 | 14 | 95% |
| 映射查表 | 2 | 9 | 100% |
| 清理與合併 | 2 | 10 | 90% |
| **合計** | **15** | **126** | **97%** |
