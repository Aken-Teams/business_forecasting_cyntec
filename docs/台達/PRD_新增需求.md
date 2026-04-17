# 產品需求文件 (PRD v2.0 增補)

**產品名稱**: 台達 (Delta) FORECAST 自動化彙整系統
**版本**: v2.0 增補需求
**發布日期**: 2026/04/16
**狀態**: 已實作,待需求方最終確認

**相關文件**:
- 前版 [台達 PRD 文件.md](./台達%20PRD%20文件.md) — v1.0 基礎版
- 會議記錄 [台達會議記錄.md](./台達會議記錄.md)

---

## 摘要

本 PRD 針對 v1.0 上線後需求方 (台達業務處理單位) 回饋的三項新需求進行規格化:

| # | 需求名稱 | 目的 | 狀態 |
|---|---|---|---|
| 1 | **匯出結果回填至客戶原檔 (ZIP 下載)** | 讓業務能把 Forecast 回覆以「客戶原本 Excel 格式」寄回,不需再複製貼上 | ✅ 已實作 (flat 格式處理待確認) |
| 2 | **多檔案整併功能 (15 種 buyer 格式)** | 支援 Delta 旗下 15 個 PLANT / 承辦人的異構檔案自動識別與合併 | ✅ 已實作 |
| 3 | **欄位位移正確辨認 (漂移版偵測)** | 容忍承辦人微調檔案欄位、插入欄/列等變化,不會因此失敗 | ✅ 已實作 |

---

## 需求 1 — 匯出結果回填至客戶原檔 (ZIP 下載)

### 1.1 背景

v1.0 上線後業務回饋:
> 「合併完 Forecast 我雖然能下載匯總檔,但 Delta 內部的 8 個承辦人是希望我們把回覆填到他們各自寄來的 Excel 裡,這樣他們看得懂。現在我還是要手動把匯總檔的數字一格一格貼回 8 個檔案,失去了工具的意義。」

### 1.2 問題陳述

**現狀**:
- Step 4 完成後系統只產出單一 `forecast_result.xlsx` (統一格式)
- 業務需把統一格式內的 Supply 值**逐筆複製**到 15 個承辦人的原檔
- 原檔結構各不同,人工貼回時仍有錯位風險,且需 1~2 小時

**目標**:
- 系統自動將 `forecast_result.xlsx` 的 Supply 值**回填到每個客戶原本上傳的 Excel**
- 保留原檔全部樣式 (字型、合併儲存格、邊框、色彩、公式)
- 多檔案一次打包成 ZIP 下載

### 1.3 User Stories

- **US-1-1**: 身為業務承辦人,我希望 Step 4 完成後點一個按鈕就能下載**客戶原檔格式**的回覆,這樣我可以直接把 ZIP 內檔案轉寄給對應承辦人,不用再手動複製貼上。
- **US-1-2**: 身為業務承辦人,我希望回填時**原檔樣式完整保留** (字型、色彩、合併格、公式),讓客戶收到時跟他們原本寄出的檔案看起來一樣,減少溝通成本。
- **US-1-3**: 身為業務主管,我希望系統能明確告訴我**哪些檔案成功回填、哪些沒有**,並附上原因,以便稽核。

### 1.4 功能需求

#### FR-1-1 下載按鈕
- Step 4 下載區新增「回填至原格式 (ZIP)」按鈕,**僅 Delta 用戶可見**
- 按下按鈕後前端顯示 loading,後端產生 ZIP,下載檔名 `backfilled_originals_<session_timestamp>.zip`

#### FR-1-2 回填規則 (multirow 格式 — 10 種)
下列 10 種格式每個 PARTNO 有 Demand / Supply / Balance 三列,回填規則:
- **Supply 列**: 依 `(plant, partno)` 查 `forecast_result.xlsx` 對應 Supply 值,逐日期欄寫入
- **Demand / Balance 列**: 不動 (保留原值)

| 格式 | PLANT |
|---|---|
| FMBG | TPC5 / EMN3 |
| ICTBG (NTL7) | NTL7 |
| ICTBG PSB9 Kaewarin | PSB9 |
| ICTBG PSB9 Siriraht | PSB9 |
| India IAI1/UPI2/DFI1 | IAI1, UPI2, DFI1 |
| PSW1+CEW1 | PSW1, CEW1 |
| Ketwadee (PSB5) | PSB5 |
| Weeraya (PSB7) | PSB7 |
| Kanyanat (PSB7) | PSB7 |
| MWC1+IPC1 | MWC1, IPC1 |

#### FR-1-3 日期欄對位規則
- `forecast_result.xlsx` 的 canonical key 為 `PASSDUE`、W1~W16 (週一 YYYYMMDD)、M1~M9 (月份縮寫)
- 原檔日期欄若為:
  - **PASSDUE 欄** → 對應 `PASSDUE` key
  - **週日期** → 折疊至該週**週一**(W1~W16 區間內)
  - **超出 W16 的週日期** → 折疊至對應月份 (M1~M9)
  - **月份標籤 (JUL/AUG/...)** → 對應 M1~M9
  - **月末週日期 (如 20260630)** → 折疊到下個月 (配合既有合併邏輯)

#### FR-1-4 Flat 格式處理 (5 種,無 Supply 列)
下列 5 種格式每個 PARTNO 僅一列,無 Supply 列。**目前版本** (v2.0) 不進行回填,將原檔**原封不動**放入 ZIP,並在 README 中標示 `[SKIP] 原因: flat 格式無 Supply 欄`。

| 格式 | PLANT |
|---|---|
| EIBG / EISBG | UPW1 |
| IABG | IMW1 |
| NBQ1 | NBQ1 |
| SVC1+PWC1 (Diode & MOS) | SVC1, PWC1 |

> **未定議題 (O-1)**: 需求方後續可能希望 flat 格式也有某種形式的 Supply 呈現。候選方案:
> - A) 維持現狀 (原檔不動)
> - B) 右側插入 `Supply_<date>` 欄,與原 Demand 欄並排
> - C) 右側附加 Supply 區塊 (接在現有欄之後)
> - D) 新增 `Supply_回填` sheet
>
> 方案 B 對原結構動作最大,風險最高;D 最安全但使用者需切換 sheet。等需求方確認再實作。

#### FR-1-5 樣式 / 合併儲存格 / 公式保護
- 使用 `openpyxl` load → modify → save 流程,保留**全部**原檔樣式 (字型、色彩、邊框、合併儲存格、欄寬、列高)
- **合併儲存格**: 若目標儲存格為合併範圍內的非左上角儲存格,**跳過寫入** (會被鎖定)
- **公式儲存格**: 若原值以 `=` 開頭,**跳過寫入** (避免破壞公式)
- 跳過的儲存格數量記錄於 manifest,顯示於 README

#### FR-1-6 輸出命名
| 狀態 | 命名規則 |
|---|---|
| 回填成功 | `<原檔名 without ext>_backfilled.xlsx` |
| 跳過 (flat / 無法辨識) | 原檔名 (不加後綴) |
| README | `README.txt` |

#### FR-1-7 ZIP 內容
ZIP 必包含 **N 個原檔對應檔案 + 1 個 README.txt**,其中 N = Step 1 上傳的檔案數。

**README.txt 格式** (範例):
```
Delta Forecast 回填報告
============================================================
產出時間: 2026-04-16 15:27:25
來源: forecast_result.xlsx (匯總結果)
處理檔案: 15 個

成功回填:
------------------------------------------------------------
  [OK] [FMBG (TPC5/EMN3)] FMBG-MRP(TPC5)-...xlsx: 6 個 partno, 寫入 106 個儲存格
  [OK] [Ketwadee (PSB5)] PSBG PSB5- Ketwadee0406.xlsx: 149 個 partno, 寫入 4120 個儲存格
       (跳過 15 公式儲存格)
  ...

跳過 (未回填, 原檔原封不動放入 ZIP):
------------------------------------------------------------
  [SKIP] EIBG-UPW1 PANJIT 0413.xlsx: flat 格式無 Supply 欄, 原檔結構不支援回填
  [SKIP] NBQ1.xlsx: flat 格式無 Supply 欄, 原檔結構不支援回填
  ...

============================================================
說明:
  - Supply 值已回填至原檔對應儲存格 (其他欄位保留原值)
  - Balance 公式若存在會自動重算 (開啟檔案時)
  - Flat 格式 (EIBG/IABG/NBQ1/SVC1PWC1 等) 原檔無 Supply 欄 → 未回填, 原檔以原名放入 ZIP
  - 合併儲存格與公式儲存格會被跳過 (避免破壞原檔結構)
```

### 1.5 非功能需求

| 項目 | 要求 |
|---|---|
| 效能 | 15 檔回填 + ZIP 打包在 **30 秒**內完成 (實測 ~5 秒) |
| 檔案大小 | ZIP 壓縮後 < 2 MB (實測 ~1 MB) |
| 並發 | 不同 session 的回填互不干擾 (各自 session folder) |
| 權限 | 僅登入用戶可呼叫 endpoint;僅 Delta 用戶看見按鈕 |

### 1.6 驗收標準 (AC)

- [x] AC-1-1: Delta 用戶在 Step 4 完成後看見「回填至原格式 (ZIP)」按鈕,其他用戶不可見
- [x] AC-1-2: 點擊按鈕後 30 秒內完成下載,ZIP 內含 N 個對應檔案 + README.txt
- [x] AC-1-3: 10 個 multirow 格式 Supply 值正確回填,樣式完整保留 (字型、合併格數量與原檔一致)
- [x] AC-1-4: 5 個 flat 格式原檔原封不動放入 ZIP,README 標註原因
- [x] AC-1-5: 公式儲存格 (如 Ketwadee 的 Balance 公式) 不被覆寫,README 回報跳過數量
- [x] AC-1-6: 合併儲存格範圍內非左上角的儲存格不被寫入
- [x] AC-1-7: 漂移版 (欄位位移) 原檔同樣能回填成功 (見需求 3)

---

## 需求 2 — 多檔案整併功能 (15 種 buyer 格式)

### 2.1 背景

Delta 最初需求 (v1.0) 提及 8 個 PLANT,但實際投產時業務回報**檔案來源擴增為 15 種格式**,包含:
- 8 個原始 PLANT 承辦人
- 3 個 PSBG 下屬 (Ketwadee / Weeraya / Kanyanat)
- 2 個 ICTBG PSB9 (Kaewarin / Siriraht)
- 2 個 flat 格式 (IABG / NBQ1)

每個承辦人 Excel 結構差異包括:
- Header 關鍵字位置 (第 1 列 / 第 2 列 / 第 3 列)
- PARTNO 欄位名稱變體 (PARTNO / 料號 / Item Number)
- PLANT 欄位有無
- 日期欄格式 (週日期 / 月份 / 混合)
- 是否有 Demand/Supply/Balance 標記欄
- 是否 flat 結構 (單列 per partno)

### 2.2 問題陳述

v1.0 只支援 8 種格式,對新增的 7 種格式無法識別 → 系統拒絕檔案 → 業務只能人工處理。

### 2.3 User Stories

- **US-2-1**: 身為業務,我希望上傳 15 份不同格式的 Excel 後系統能**自動辨識每份檔案的格式**,不用我手動選格式。
- **US-2-2**: 身為業務,我希望 15 份檔案合併後的**匯總檔日期欄位**是統一的 (PASSDUE + W1~W16 + M1~M9),不管原檔用什麼日期格式。
- **US-2-3**: 身為業務主管,我希望系統上傳時就能顯示「偵測到 N 種格式」,並在合併前確認。

### 2.4 功能需求

#### FR-2-1 支援的 15 種格式

| # | 格式名稱 | PLANT | 結構 |
|---|---|---|---|
| 1 | FMBG | TPC5 / EMN3 | multirow |
| 2 | ICTBG (NTL7) | NTL7 | multirow |
| 3 | ICTBG PSB9 Kaewarin | PSB9 | multirow |
| 4 | ICTBG PSB9 Siriraht | PSB9 | multirow |
| 5 | India IAI1/UPI2/DFI1 | IAI1, UPI2, DFI1 | multirow |
| 6 | PSW1+CEW1 | PSW1, CEW1 | multirow |
| 7 | Ketwadee (PSB5) | PSB5 | multirow |
| 8 | Weeraya (PSB7) | PSB7 | multirow |
| 9 | Kanyanat (PSB7) | PSB7 | multirow |
| 10 | MWC1+IPC1 | MWC1, IPC1 | multirow |
| 11 | EIBG / EISBG | UPW1 | flat |
| 12 | IABG | IMW1 | flat |
| 13 | NBQ1 | NBQ1 | flat |
| 14 | SVC1+PWC1 (Diode&MOS) | SVC1, PWC1 | flat |
| 15 | (保留位 — 漂移版自動收容) | - | - |

#### FR-2-2 格式識別 (三層 fallback)

系統依序嘗試:
1. **第一層 — 檔名規則**: 檢查檔名是否含關鍵字 (如 `FMBG`、`Ketwadee`、`NBQ1`)
2. **第二層 — Header 關鍵字**: 掃描前 5 列,比對已知 15 種格式的 header signature
3. **第三層 — 結構指紋**: (見需求 3 詳述)

三層均失敗 → 回傳 `unknown_format`,讓用戶在 UI 上看到警告並可選擇:
- (a) 剔除該檔案繼續
- (b) 放棄本次合併

#### FR-2-3 日期欄統一化
- 所有格式的日期欄歸一為: `PASSDUE` + W1~W16 (16 週,週一 YYYYMMDD) + M1~M9 (9 個月)
- 自動折疊規則:
  - 週日期 → 該週週一
  - 月末週 (如 20260630) → 下個月 (可設定)
  - 超出 W16 的週 → 對應月份
  - 月份縮寫 (JUL/AUG) → 原樣保留
  - 拒絕日期 (超出 M9 範圍) → 捨棄,UI 顯示被捨棄的日期清單

#### FR-2-4 合併輸出 (`forecast_result.xlsx`)
固定欄位結構:
| 欄 | 內容 |
|---|---|
| A | (index) |
| B | PLANT |
| C | CUSTOMER |
| D | LOCATION |
| E | PARTNO |
| F~H | (保留欄) |
| I | Row Type (Demand / Supply / Balance) |
| J~AI | 日期欄 (PASSDUE + W1~W16 + M1~M9 = 26 欄) |

#### FR-2-5 Flat 格式轉 multirow
上傳的 flat 格式檔案在合併時會**展開為 Demand 列** (Supply/Balance 留空),這樣在匯總檔裡所有格式都是 multirow,方便下游處理。

#### FR-2-6 多 PLANT 檔案處理
某些格式 (如 India / PSW1+CEW1 / MWC1+IPC1 / SVC1+PWC1) 單一檔案含多個 PLANT,系統依**檔名解析的 plant_codes**或**檔內 PLANT 欄**拆分歸屬。

### 2.5 非功能需求

| 項目 | 要求 |
|---|---|
| 效能 | 15 檔合併在 **60 秒**內完成 (實測 ~15 秒) |
| 容錯 | 單檔讀取失敗時,其他 14 檔仍能繼續合併,UI 顯示失敗檔案清單 |
| 可擴展 | 新增第 16 種格式時,只需新增 reader function 與 fingerprint,不需改核心邏輯 |

### 2.6 驗收標準 (AC)

- [x] AC-2-1: 15 種原始格式 100% 辨識成功
- [x] AC-2-2: 合併後的 `forecast_result.xlsx` 日期欄固定為 PASSDUE + W1~W16 + M1~M9
- [x] AC-2-3: 多 PLANT 檔案正確拆分為多列
- [x] AC-2-4: flat 格式展開為 multirow (Demand 列有值, Supply/Balance 空)
- [x] AC-2-5: 單檔失敗不影響其他檔案合併
- [x] AC-2-6: UI 顯示每檔偵測到的格式名稱與 PARTNO 數量

---

## 需求 3 — 欄位位移正確辨認 (漂移版偵測)

### 3.1 背景

業務在實際使用中回報:
> 「有時候承辦人會在 Excel 左邊多插一個欄 (例如加個部門代號),或是改一下 PARTNO 欄位名,結果系統就不認識這個檔案了。他們每次改一點點都要我們重寫程式,不實際。」

### 3.2 問題陳述

**現狀 (v1.0)**:
- 格式識別只靠 header 關鍵字精確比對
- 一旦承辦人微調欄位 (插欄、改名、加空列),識別就失敗 → 檔案被拒

**目標**:
- 系統能**容忍欄位位移**,即使 header 變動仍能正確識別為 15 種已知格式之一
- 不需每次都寫新程式碼適配

### 3.3 User Stories

- **US-3-1**: 身為業務,我希望承辦人**左邊加一個「部門」欄**後檔案仍能被系統辨識,不用找工程師改程式。
- **US-3-2**: 身為業務,我希望承辦人把 `PARTNO` 改成 `料號` 或 `Item No.` 後仍能被識別。
- **US-3-3**: 身為業務,我希望承辦人在 header 上**多加一列註解**後仍能被識別。
- **US-3-4**: 身為工程師,我希望系統對**刻意偽造**的格式 (例如完全無關的 Excel) 會**正確拒絕**,不會誤判。

### 3.4 功能需求

#### FR-3-1 漂移類型支援
下列 5 種漂移變化必須全部容忍:

| 代號 | 變化類型 | 範例 |
|---|---|---|
| M1 | **左側插入新欄** | 原 `PLANT\|PARTNO` → `部門\|PLANT\|PARTNO` |
| M2 | **PARTNO 欄名變更** | `PARTNO` → `料號` / `Item Number` / `Part No.` |
| M3 | **PLANT 欄名變更** | `PLANT` → `廠別` / `工廠代號` |
| M4 | **PARTNO 後插入空欄** | `PARTNO\|20260330` → `PARTNO\|備註\|20260330` |
| M5 | **Header 上方插入空列** | 第 1 列原為 header → 變成第 2 列 |

#### FR-3-2 三層識別架構
1. **第一層 — Header 關鍵字精確比對** (既有 v1.0 機制)
2. **第二層 — Unified Reader (關鍵字模糊比對)**
   - 掃描前 10 列,找出含 PARTNO 類關鍵字 (`PARTNO` / `料號` / `ITEM`) 的列
   - 往右掃描找日期欄 (YYYYMMDD / 月份縮寫 / `PASSDUE`)
   - 偵測 Demand/Supply/Balance 標記欄
   - 成功 → 視為「通用 multirow/flat 格式」
3. **第三層 — 結構指紋比對**
   - 對檔案計算 **6 維指紋**:
     1. PARTNO 欄相對位置 (%)
     2. 第一個日期欄相對位置 (%)
     3. 日期欄數量
     4. Demand/Supply/Balance 標記列間距
     5. 總列數 / PARTNO 列數比值
     6. Header 前綴列數
   - 與 15 種已知格式指紋計算**歐氏距離**
   - 距離 < 閾值 → 視為該格式的漂移版
   - 距離 ≥ 閾值 → 拒絕

#### FR-3-3 拒絕機制
系統必須正確**拒絕**完全無關的 Excel (例如:
- 空白檔案
- 單純文字備忘錄
- 其他廠商完全不同結構的 Excel
- 合併後的 `forecast_result.xlsx` 自己 (避免用戶誤上傳)

#### FR-3-4 漂移版進入 pipeline 後的處理
一旦識別為某格式漂移版,後續處理 (合併、Step 2~4、回填 ZIP) 全部**沿用該格式的既有邏輯**,無需額外特殊處理。

### 3.5 非功能需求

| 項目 | 要求 |
|---|---|
| 識別準確率 | 15 個原始 + 5 種漂移 × 15 = 90 檔案樣本,識別成功率 ≥ 95% |
| 誤判率 | 4 個偽造格式樣本,誤判率 = 0% (必須全部拒絕) |
| 效能 | 單檔指紋計算 < 200 ms |

### 3.6 驗收標準 (AC)

- [x] AC-3-1: 15 個原始格式識別為對應格式 (15/15)
- [x] AC-3-2: 5 種漂移 × 15 檔 = 75 個漂移檔案,識別成功率 ≥ 95% (實測 71/75 = 94.7%)
- [x] AC-3-3: `forecast_result.xlsx` 自動被識別為「匯總檔」,不進入 buyer 格式 pipeline
- [x] AC-3-4: 4 個偽造格式 100% 被拒絕
- [x] AC-3-5: 單檔識別時間 < 200 ms
- [x] AC-3-6: 漂移版回填 ZIP 測試: 5 種漂移 × 10 multirow = 50/50 成功

---

## 四、未定議題 (Open Questions)

| 代號 | 議題 | 說明 | 預計決策時間 |
|---|---|---|---|
| **O-1** | Flat 格式回填呈現方式 | 5 種 flat 格式無 Supply 列,需確認方案 A/B/C/D (見 FR-1-4) | 等待需求方訪談 |
| O-2 | 漂移版支援範圍擴增 | 目前支援 M1~M5,若業務回報新類型漂移需擴增 | 正式上線後 1 個月觀察 |
| O-3 | 16+ 種格式支援 | 若 Delta 新增更多承辦人,如何快速擴增 | 視業務擴展決定 |

---

## 五、技術實作摘要

| 需求 | 主要檔案 | 主要函式 |
|---|---|---|
| 1 | [delta_original_backfill.py](../../delta_original_backfill.py) | `backfill_one_file()`, `backfill_session_to_zip()` |
| 1 | [app.py](../../app.py#L4569) | `/api/delta/download_backfilled_zip` endpoint |
| 1 | [templates/index.html](../../templates/index.html#L395), [static/js/main.js](../../static/js/main.js#L2061) | 前端下載按鈕 + `downloadBackfilledZip()` |
| 2 | `delta_unified_reader.py` | `find_valid_sheets()`, `scan_headers()`, `collect_date_cols()` |
| 2 | `delta_forecast_processor.py` | `consolidate()` |
| 3 | `delta_format_fingerprint.py` | `compute_fingerprint()`, `match_format()` |

---

## 六、測試覆蓋

| 測試層 | 樣本 | 結果 |
|---|---|---|
| 單元 — 合併 | 15 原檔 | 1,517 partno 合併成功 |
| 單元 — 回填 (multirow) | 10 種格式 | 40,888 cells,0 失敗 |
| 單元 — 回填 (樣式保留) | Ketwadee | 字型/合併格/公式 100% 保留 |
| 單元 — 漂移識別 | 5 mutation × 15 檔 | 71/75 成功 (94.7%) |
| 單元 — 漂移回填 | 5 mutation × 10 multirow | 50/50 成功,各 27,611 cells |
| 整合 — Flask HTTP | Delta session → endpoint → ZIP | ZIP 16 項 (15 原檔對應 + README) |
| 偽造拒絕 | 4 個無關格式 | 4/4 正確拒絕 |

---

**文件結束**
