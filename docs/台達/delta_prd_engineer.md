# 產品需求文件 (PRD) — 工程師版

**產品名稱**: 台達 (Delta) FORECAST 自動化彙整系統
**版本**: v1.0
**發布日期**: 2026/04/10
**專案代碼**: `delta_forecast_processor` + `delta_forecast_step4`
**Repo**: `business_forecasting_cyntec`

---

## 一、背景

台達 (Delta) 為既有 `business_forecasting_cyntec` 系統新增之客戶模組。既有系統已支援廣達 (Quanta)、和碩 (Pegatron)，台達因來源檔案數量多 (8 份) 且格式差異大，需獨立開發 reader 與彙整流程，但仍共用 Flask 框架、DB 與 UI 基座。

---

## 二、痛點分析 (工程視角)

### 痛點 1 — 多格式 reader 難以維護
- 8 份檔案日期欄位數量不同 (22~45)
- PASSDUE、月份欄混合週日期欄
- 多 PLANT 情境：某份檔案的料號要依內容判斷 PLANT (India IAI1/UPI2/DFI1、PSW1+CEW1 等)
- 初期若每個 reader 各自處理日期累加，會有維護地獄

### 痛點 2 — 日期折疊邏輯容易寫錯
- W1 起點該用「今天的週一」還是「來源檔最早的週一」？
- PASSDUE 是獨立欄還是月份？
- 超出範圍的日期應丟棄還是折入最近月份？
- 若規則不明確，下游結果隨時間漂移

### 痛點 3 — 多對一映射易被去重
- 固定 26 欄代表 source 可能有 5 個週日期對應到同一個月份 target
- 若 `_build_date_col_map` 用 `seen_keys` 去重，會遺失 4 筆資料
- 需要在 row 層級累加而非 col 層級

### 痛點 4 — Windows 環境踩坑
- 檔名含非破壞空白 (`\xa0`)，無法用 raw string 硬寫
- cp950 console 無法輸出 emoji，需設定 `PYTHONIOENCODING=utf-8`
- 中文路徑需注意 encoding

### 痛點 5 — 測試資料難以完整覆蓋
- 8 份真實檔案 + ERP 總共 ~1 萬行資料
- 單元測試僅能覆蓋局部
- 需要端到端測試驗證全流程

---

## 三、技術目標

| 目標 | 實作產物 |
|------|---------|
| 統一 8 檔 reader 架構 | `_read_row_dates()` helper |
| 固定 26 欄輸出 | `_sort_date_cols()` 方案二 |
| 多對一累加 | 移除 `_build_date_col_map` 的 `seen_keys` |
| 端到端可測 | `tmp/delta_e2e/test_full_pipeline.py` |
| 跨客戶隔離 | DB `user_id=7` |

---

## 四、User Stories (工程視角)

### Story 1 — 可擴充的 Reader 架構
**As a** 維運工程師
**I want** 新增一個來源 PLANT 的 reader 只需要複製現有 reader 並改欄位對應
**So that** 客戶新增來源時不需要重寫整個流程

**驗收條件**:
- ✅ 每個 reader 統一使用 `_read_row_dates(row, date_col_map)` 讀取日期欄
- ✅ reader 只需負責：找到料號欄、找到日期 header、判斷 PLANT
- ✅ 新增 reader 不需修改 `consolidate()` 主流程 (只需註冊)
- ✅ 日期折疊邏輯完全由 `_sort_date_cols()` 集中處理

**相關檔案**: [delta_forecast_processor.py](../../delta_forecast_processor.py)

---

### Story 2 — 固定 26 欄輸出且 W1 可動態
**As a** 後端工程師
**I want** `_sort_date_cols()` 產出永遠是 PASSDUE + 16 週 + 9 月 的 26 欄
**So that** 下游 Step 4 可安全假設固定欄位結構

**驗收條件**:
- ✅ `final_cols` 長度 == 26
- ✅ `final_cols[0]` == `'PASSDUE'`
- ✅ `final_cols[1:17]` 為 16 個連續 Monday 日期字串 (YYYYMMDD)
- ✅ `final_cols[17:26]` 為 9 個月份標籤 (from `MONTH_NAMES`)
- ✅ W1 起點 = 來源檔最早的週日期對齊到 Monday
- ✅ 若來源無週日期，fallback 為今天的 Monday
- ✅ `conversions` 字典涵蓋所有來源日期的歸位

**程式碼骨架**:
```python
def _sort_date_cols(dates, anchor_date=None):
    # 1. 找來源最早週一
    weekly_source = [datetime.strptime(d, '%Y%m%d')
                     for d in dates if isinstance(d, str) and d.isdigit() and len(d) == 8]
    first_monday = (min(weekly_source) if weekly_source else (anchor_date or datetime.now()))
    first_monday -= timedelta(days=first_monday.weekday())
    first_monday = first_monday.replace(hour=0, minute=0, second=0, microsecond=0)

    # 2. 生成 16 週 + 9 月
    weekly_keys = [(first_monday + timedelta(weeks=i)).strftime('%Y%m%d') for i in range(16)]
    # ... (month labels from W17)

    # 3. 建立 conversions 折疊表
    return ['PASSDUE'] + weekly_keys + monthly_keys, conversions
```

---

### Story 3 — PASSDUE 語意嚴格
**As a** QA
**I want** PASSDUE 欄只接收來源檔明確標記為 PASSDUE 的資料
**So that** 不會因為早於 W1 的日期被誤折入

**驗收條件**:
- ✅ 源檔欄位為 `"PASSDUE"`, `"PAST DUE"`, `"passdue"` → 歸入 PASSDUE
- ✅ 源檔日期早於 `first_monday` → **丟棄** (不進 PASSDUE)
- ✅ 由於 `first_monday = min(weekly_source)`，理論上不應出現早於 W1 的日期
- ✅ 若出現仍進入 `rejected` 列表並 print warning
- ❌ **禁止** `conversions[d] = 'PASSDUE'` 於日期折疊邏輯中出現

**反例 (初版 bug)**:
```python
# ❌ 錯誤：會把 20260330 誤折入 PASSDUE
if dt < first_monday:
    conversions[d] = 'PASSDUE'
```

**正例**:
```python
# ✅ 正確：丟棄並警告
if dt < first_monday:
    conversions[d] = None
    rejected.append(d)
```

---

### Story 4 — Row-level 累加
**As a** 後端工程師
**I want** 多個 source 欄位指向同一個 target key 時在 row 層級累加
**So that** 9/7, 9/14, 9/21, 9/28 四個週欄位資料都能正確加到 SEP

**驗收條件**:
- ✅ `_build_date_col_map()` 移除 `seen_keys` 去重
- ✅ `date_col_map` 允許多個 col_idx → 同一個 date_key
- ✅ `_read_row_dates()` 使用 `data[key] = data.get(key, 0) + v_num` 累加
- ✅ 輸出的匯總檔中，月份欄數值 = 該月所有週日期的總和

**實作**:
```python
def _read_row_dates(row, date_col_map):
    data = {}
    for col_idx, date_key in date_col_map.items():
        v = row[col_idx - 1].value
        if v is None or v == '':
            continue
        try:
            v_num = float(v)
        except (ValueError, TypeError):
            continue
        data[date_key] = data.get(date_key, 0) + v_num
    return data
```

---

### Story 5 — Step 4 Transit/ERP 回填
**As a** 後端工程師
**I want** `process_delta_forecast()` 能依 (客戶簡稱, 送貨地點, 料號) 精準回填
**So that** 最終結果與 ERP 原始資料一致

**驗收條件**:
- ✅ Transit 填入 OTW QTY 列，依 key lookup
- ✅ ERP 淨需求填入 Demand 列，依排程出貨日期歸入對應週/月
- ✅ 若日期早於 W1 或晚於 M9 → 跳過並記入 `erp_skipped`
- ✅ 返回統計字典: `{transit_filled, transit_skipped, transit_matched_rows, erp_filled, erp_skipped, erp_matched_rows}`
- ✅ 保留 Forecast 原始格式 (客戶收到的檔案外觀不變)

**相關檔案**: [delta_forecast_step4.py](../../delta_forecast_step4.py)

---

### Story 6 — 客戶隔離
**As a** 系統架構師
**I want** 台達資料與廣達/和碩資料完全隔離
**So that** 不同客戶的映射表互不干擾

**驗收條件**:
- ✅ Delta 使用 `user_id=7`
- ✅ `customer_mappings` 所有查詢帶 `WHERE user_id=%s`
- ✅ 登入不同客戶後 session 綁定對應 `user_id`
- ✅ 映射 UI 只顯示當前登入使用者的映射

---

### Story 7 — 端到端可測試
**As a** QA 工程師
**I want** 一個腳本可以跑完整個 Delta pipeline
**So that** 發佈前能快速驗證

**驗收條件**:
- ✅ `tmp/delta_e2e/test_full_pipeline.py` 可獨立執行
- ✅ 覆蓋 Step 1 → Step 3 → Step 4 完整流程
- ✅ 使用真實 8 份 Forecast + 真實 ERP 檔案
- ✅ Transit 若無則用 fake generator 產生
- ✅ 執行結束輸出統計與驗證結果
- ✅ 預期結果: 1204 料號 / 26 欄 / Demand 有值 / PASSDUE 有值

**執行方式**:
```bash
cd /d/github/business_forecasting_cyntec
PYTHONIOENCODING=utf-8 python tmp/delta_e2e/test_full_pipeline.py
```

**相關檔案**: [tmp/delta_e2e/test_full_pipeline.py](../../tmp/delta_e2e/test_full_pipeline.py)

---

### Story 8 — Windows 環境相容
**As a** 部署工程師
**I want** 系統在 Windows 11 環境下能正常處理中文檔名與 emoji 輸出
**So that** 現場可直接在辦公電腦執行

**驗收條件**:
- ✅ 支援檔名含非破壞空白 (`\xa0`)
- ✅ console 能輸出 emoji (需設 `PYTHONIOENCODING=utf-8`)
- ✅ 中文路徑可正常讀寫
- ✅ 啟動指令統一: `PYTHONIOENCODING=utf-8 python app.py`

**已知陷阱**:
```python
# ❌ 錯誤: SyntaxError
filename = r'PSBG\xa0PSB5-\xa0Ketwadee.xlsx'

# ✅ 正確
NBSP = '\xa0'
filename = f'PSBG{NBSP}PSB5-{NBSP}Ketwadee.xlsx'
```

---

## 五、技術架構

### 5.1 技術棧

| 類別 | 技術 |
|------|------|
| 後端框架 | Flask (debug mode, port 12058) |
| Excel 處理 | `openpyxl` (主) + `pandas` (ERP) |
| 資料庫 | MySQL (via `pymysql`) |
| 前端 | HTML5 + CSS3 + Vanilla JS |
| Python 版本 | 3.8+ |

### 5.2 核心檔案

| 檔案 | 職責 |
|------|------|
| [app.py](../../app.py) | Flask 主應用 / `/run_forecast` Delta 分支 |
| [delta_forecast_processor.py](../../delta_forecast_processor.py) | 8 檔合併 + 方案二匯總 |
| [delta_forecast_step4.py](../../delta_forecast_step4.py) | Transit + ERP 回填 |
| [database.py](../../database.py) | MySQL / 使用者 / 映射表 |

### 5.3 Pipeline

```
[8 Forecast]  ─┐
               ├─► Step 1: consolidate() ─► consolidated.xlsx (26 cols)
[DB mapping]  ─┤
[ERP file]    ─┼─► Step 2/3: ERP map + C/D fill
[Transit]     ─┘                │
                                ▼
                     Step 4: process_delta_forecast()
                                │
                                ▼
                      forecast_result.xlsx
```

---

## 六、範圍外 (Out of Scope)

- ❌ 非同步處理 (Celery/Redis)
- ❌ 多使用者同時處理同一份檔案
- ❌ ERP 直接 API 串接
- ❌ 自動 Email 寄送結果
- ❌ 行動裝置 App
- ❌ 歷史趨勢分析

---

## 七、技術驗收 (已完成)

### 7.1 最新測試結果 (2026/04/10)

```
Step 1: 1204 料號合併成功
        W1=20260330, W16=20260713
        月份: JUL~MAR
        丟棄: ['APR','MAY','JUN','20270430','20270531','20270630']
        折疊:
          SEP ← [20260907,20260914,20260921,20260928,20260930]
          AUG ← [20260803,20260810,20260817,20260824,20260831,20270831]
          ...
Step 3: ERP 1755/2633 匹配
        Forecast C/D 3612/3612 填入
Step 4: Transit 45 cells 填入
        ERP 4188 cells 填入
最終驗證: 3612 列 / 1052 Demand 有值 / 316 列 PASSDUE 有值
✅ 全部通過
```

### 7.2 性能指標

| 項目 | 實測值 |
|------|--------|
| Step 1 (8 檔合併) | < 10 秒 |
| Step 3 (ERP 映射) | < 5 秒 |
| Step 4 (回填) | < 5 秒 |
| 全流程 | < 30 秒 |

---

## 八、已知問題與技術債

| 項目 | 嚴重度 | 備註 |
|------|--------|------|
| 異常日期容錯 (如 `20591231`) | 低 | 目前仍會依月份折疊，可加入合理性檢查 |
| 硬編 `16`/`9` 週月數 | 低 | 未來若客戶需求變動需提取為參數 |
| Mapping CRUD UI | 中 | 目前需直接改 SQL |
| Reader 新增流程 | 低 | 需改 `consolidate()` 主流程，可考慮 registry pattern |
| 錯誤 log 集中化 | 中 | 目前散在 print，應改為 logger |

---

## 九、優先順序

| Story | 優先級 | 狀態 |
|-------|--------|------|
| Story 1 — Reader 架構 | P0 | ✅ 完成 |
| Story 2 — 固定 26 欄 | P0 | ✅ 完成 |
| Story 3 — PASSDUE 語意 | P0 | ✅ 完成 (修正一次) |
| Story 4 — Row-level 累加 | P0 | ✅ 完成 |
| Story 5 — Step 4 回填 | P0 | ✅ 完成 |
| Story 6 — 客戶隔離 | P0 | ✅ 完成 |
| Story 7 — E2E 測試 | P1 | ✅ 完成 |
| Story 8 — Windows 相容 | P1 | ✅ 完成 |

---

## 十、相關文件

- [README.md](../../README.md) — 系統總覽
- [台達 PRD 文件.md](台達%20PRD%20文件.md) — 客戶版 PRD (痛點 + User Stories)
- [台達會議記錄.md](台達會議記錄.md) — 需求會議紀錄
- [DEPLOYMENT_GUIDE.md](../../DEPLOYMENT_GUIDE.md) — 部署指南

---

*本文件為工程師版 PRD，聚焦於技術痛點、User Stories 與實作驗收條件。*
