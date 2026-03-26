# 光寶 Forecast 系統 - BDD 行為規格

> Behavior-Driven Development 行為文件
> 以 Given-When-Then 格式描述系統行為，從使用者與業務邏輯角度出發

---

## 目錄

1. [Feature: 客戶映射配置](#feature-1-客戶映射配置)
2. [Feature: Forecast 檔案清理](#feature-2-forecast-檔案清理)
3. [Feature: ERP 資料映射](#feature-3-erp-資料映射)
4. [Feature: Transit 在途映射](#feature-4-transit-在途映射)
5. [Feature: ERP 日期計算與填入](#feature-5-erp-日期計算與填入)
6. [Feature: Transit 日期填入](#feature-6-transit-日期填入)
7. [Feature: Forecast 多檔合併](#feature-7-forecast-多檔合併)
8. [Feature: 日期欄位查找策略](#feature-8-日期欄位查找策略)

---

## Feature 1: 客戶映射配置

```gherkin
Feature: 客戶映射配置管理
  作為光寶的 Forecast 管理人員
  我需要配置客戶與 Plant 的映射關係
  以便 ERP/Transit 資料能正確填入對應的 Forecast 檔案

  Background:
    Given 系統已登入光寶帳號 (user_id=6)
    And 映射配置頁面已載入

  # ── 基本 CRUD ──

  Scenario: 載入既有映射配置
    Given 資料庫已存在 21 筆映射記錄
    When 開啟映射配置頁面
    Then 顯示 21 筆記錄（分頁顯示，每頁 10 筆）
    And 第 1 頁顯示前 10 筆
    And 頁碼區顯示 "共 21 筆記錄，第 1 / 3 頁"

  Scenario: 新增映射記錄
    Given 目前有 20 筆映射記錄
    When 點擊 "+新增客戶" 按鈕
    Then 在列表末端新增一筆空白記錄
    And 自動跳轉到最後一頁
    And 聚焦在新記錄的「客戶簡稱」欄位

  Scenario: 刪除映射記錄
    Given 第 5 筆記錄為 "光-常州 / 2680"
    When 點擊該記錄的刪除按鈕
    Then 顯示確認對話框 "確定要刪除「光-常州 - 2680」嗎？"
    When 點擊「確認刪除」
    Then 該記錄從列表移除
    And 顯示通知 "已刪除「光-常州 - 2680」"

  # ── ETD/ETA 配置 ──

  Scenario: 設定 ETD 為「本週五」
    Given 正在編輯 "光-常州" 的映射
    When ETD 第一個下拉選「本週」
    And ETD 第二個下拉選「五」
    Then ETD 值組合為 "本週五"

  Scenario: 設定 ETA 為「下下下週三」
    Given 正在編輯 "光-429E)" 的映射
    When ETA 第一個下拉選「下下下週」
    And ETA 第二個下拉選「三」
    Then ETA 值組合為 "下下下週三"

  Scenario: 載入含「下下下週」的 ETA 值
    Given 資料庫中 "光-429E)" 的 ETA 為 "下下下週三"
    When 載入映射配置頁面
    Then ETA 週別下拉顯示 "下下下週"
    And ETA 星期下拉顯示 "三"

  # ── 分頁與保存 ──

  Scenario: 跨頁編輯後保存
    Given 在第 1 頁修改了 "光-常州" 的 ETA 為 "下週四"
    And 切換到第 2 頁
    And 修改了 "光-429E)" 的斷點為 "禮拜五"
    When 點擊「保存配置」
    Then "光-常州" 的 ETA "下週四" 正確保存到資料庫
    And "光-429E)" 的斷點 "禮拜五" 正確保存到資料庫

  Scenario: 保存時驗證必填欄位
    Given 有一筆記錄的「客戶簡稱」為空
    When 點擊「保存配置」
    Then 顯示錯誤通知 "客戶簡稱不能為空，請檢查後再保存"
    And 不執行保存

  Scenario: 保存時偵測重複組合
    Given 有兩筆記錄的 (客戶簡稱, 地區) 都是 ("光-常州", "2680")
    When 點擊「保存配置」
    Then 顯示錯誤通知 "存在重複的（客戶簡稱 + 地區）組合"
    And 不執行保存

  # ── 光寶專屬欄位 ──

  Scenario: 顯示光寶專屬欄位
    Given 映射記錄中至少一筆有 order_type 值
    When 渲染表格
    Then 表頭額外顯示「訂單型態」「送貨地點」「倉庫」「日期算法」欄位

  Scenario: Type 11 訂單配置
    Given 新增一筆映射記錄
    When 訂單型態選「11」
    And 送貨地點輸入「常州」
    And 日期算法選「ETA」
    Then 該記錄用於匹配 11一般訂單 的 ERP 資料

  Scenario: Type 32 訂單配置
    Given 新增一筆映射記錄
    When 訂單型態選「32」
    And 倉庫輸入「倉庫」
    And 日期算法選「ETA」
    Then 該記錄用於匹配 32HUB補貨單 的 ERP 資料
```

---

## Feature 2: Forecast 檔案清理

```gherkin
Feature: Forecast 檔案上傳與清理
  作為光寶的 Forecast 管理人員
  我需要上傳 Forecast 檔案並清除舊的 Commit 資料
  以便重新填入最新的 ERP/Transit 數據

  Background:
    Given 系統已登入光寶帳號

  Scenario: 上傳單個 Forecast 檔案
    Given 使用者上傳一個 "forecast_data.xlsx"
    And 檔案包含 "Daily+Weekly+Monthly" sheet
    When 執行清理程序
    Then 所有 Commit row 的日期欄位 (K~BY) 被清零
    And Demand row 保持原值
    And Accumulate Shortage row 保持原值
    And 產出 "cleaned_forecast.xlsx"

  Scenario: 上傳多個 Forecast 檔案
    Given 使用者上傳 23 個 "forecast_data_1.xlsx" ~ "forecast_data_23.xlsx"
    When 執行清理程序
    Then 每個檔案的 Commit row 日期欄位被清零
    And 產出 23 個 "cleaned_forecast_1.xlsx" ~ "cleaned_forecast_23.xlsx"

  Scenario: 檔案缺少必要 Sheet
    Given 上傳的檔案沒有 "Daily+Weekly+Monthly" sheet
    When 執行清理程序
    Then 回報錯誤 "Sheet 'Daily+Weekly+Monthly' not found"
```

---

## Feature 3: ERP 資料映射

```gherkin
Feature: ERP 資料映射處理
  作為系統
  我需要根據客戶映射配置將 ERP 資料補上地區、斷點、ETD/ETA 等欄位
  以便後續日期計算能正確執行

  Background:
    Given 已上傳 ERP 資料檔
    And 已配置客戶映射表

  Scenario: Type 11 一般訂單映射
    Given ERP 一筆資料:
      | 客戶簡稱 | 訂單型態    | 送貨地點 |
      | 光-常州   | 11一般訂單 | 常州     |
    And 映射表有:
      | customer_name | order_type | delivery_location | region | schedule_breakpoint | eta    | date_calc_type |
      | 光-常州        | 11         | 常州               | 2680   | 禮拜一               | 下週四 | ETA            |
    When 執行 ERP 映射
    Then 該筆 ERP 填入:
      | 客戶需求地區 | 排程出貨日期斷點 | ETA    | 日期算法 |
      | 2680         | 禮拜一           | 下週四 | ETA      |

  Scenario: Type 32 HUB 補貨單映射
    Given ERP 一筆資料:
      | 客戶簡稱  | 訂單型態      | 倉庫 |
      | 光-429E)  | 32HUB補貨單  | 倉庫 |
    And 映射表有:
      | customer_name | order_type | warehouse | region | schedule_breakpoint | eta        | date_calc_type |
      | 光-429E)       | 32         | 倉庫       | 429E   | 禮拜五               | 下下下週三 | ETA            |
    When 執行 ERP 映射
    Then 該筆 ERP 填入:
      | 客戶需求地區 | 排程出貨日期斷點 | ETA        | 日期算法 |
      | 429E         | 禮拜五           | 下下下週三 | ETA      |

  Scenario: 未匹配的 ERP 資料
    Given ERP 一筆資料的 (客戶簡稱, 送貨地點) 不存在於映射表
    When 執行 ERP 映射
    Then 該筆 ERP 的映射欄位保持空白
    And 後續處理時被跳過

  Scenario: 訂單型態前綴解析
    Given ERP 訂單型態為 "11一般訂單"
    When 解析訂單型態前綴
    Then 取得前綴 "11"
    And 使用 Type 11 查表邏輯
```

---

## Feature 4: Transit 在途映射

```gherkin
Feature: Transit 在途資料映射
  作為系統
  我需要根據 Transit 資料的送貨地點或倉庫找到對應的 Plant 地區
  以便 Transit 數量能填入正確的 Forecast 欄位

  Scenario: Type 11 Transit 映射
    Given Transit 一筆資料:
      | 訂單型態 | 送貨地點 | Ordered Item | Qty | ETA        |
      | 11       | 常州     | MAT001       | 5   | 2026-04-09 |
    And 映射表有 delivery_location="常州" → region="2680"
    When 執行 Transit 映射
    Then 該筆 Transit 填入 客戶需求地區="2680"

  Scenario: Type 32 Transit 映射
    Given Transit 一筆資料:
      | 訂單型態 | 倉庫 | Ordered Item | Qty | ETA        |
      | 32       | 倉庫 | MAT002       | 8   | 2026-04-15 |
    And 映射表有 warehouse="倉庫" → region="429E"
    When 執行 Transit 映射
    Then 該筆 Transit 填入 客戶需求地區="429E"
```

---

## Feature 5: ERP 日期計算與填入

```gherkin
Feature: ERP 日期計算與 Forecast 填入
  作為系統
  我需要根據 ERP 的排程出貨日期、斷點和 ETD/ETA 文字計算出目標日期
  然後將 淨需求×1000 的數量填入 Forecast 對應的 Commit row

  Background:
    Given 已完成 ERP 映射
    And Forecast 檔案已清理

  # ── 斷點=禮拜一 組（2680, PD00, PS00, PG00, PE00, PI00, 3560）──

  Scenario: 正常情況 - 排程日不在斷點上
    Given ERP 資料:
      | 排程出貨日期 | 斷點    | ETA    | 日期算法 | 客戶需求地區 | 客戶料號 | 淨需求 |
      | 2026-03-10   | 禮拜一 | 下週四 | ETA      | 2680         | 768290   | 150    |
    When 計算目標日期
    Then 排程=3/10(二) → week_end=3/16(一)
    And on_breakpoint=False（排程≠斷點）
    And 目標日期=3/19(四)
    And 填入值=150,000 到 3/19 對應的 Daily 欄位

  Scenario: on_breakpoint - 排程日剛好是斷點日（關鍵修復案例）
    Given ERP 資料:
      | 排程出貨日期 | 斷點    | ETA    | 日期算法 | 客戶需求地區 | 客戶料號 | 淨需求 |
      | 2026-03-30   | 禮拜一 | 下週四 | ETA      | 2680         | 768290   | 150    |
    When 計算目標日期
    Then 排程=3/30(一) → week_end=3/30(一)
    And on_breakpoint=True（排程==斷點）
    And 目標日期=4/9(四)（不是 4/2）
    And 填入值=150,000 到 4/9 對應的欄位

  Scenario: ETD 日期算法
    Given ERP 資料:
      | 排程出貨日期 | 斷點    | ETD    | ETA    | 日期算法 | 客戶需求地區 | 客戶料號 | 淨需求 |
      | 2026-03-16   | 禮拜一 | 下週一 | 下週四 | ETD      | 15K1         | MAT001   | 20     |
    When 計算目標日期
    Then 使用 ETD="下週一"（而非 ETA）
    And 目標日期=3/23(一)

  # ── 斷點=禮拜五 組（429E, 429H, 429L, 42P0, 42V0）──

  Scenario: 禮拜五斷點 - 正常情況
    Given ERP 資料:
      | 排程出貨日期 | 斷點    | ETA        | 日期算法 | 客戶需求地區 |
      | 2026-03-10   | 禮拜五 | 下下下週三 | ETA      | 429E         |
    When 計算目標日期
    Then 排程=3/10(二) → week_end=3/13(五)
    And on_breakpoint=False
    And 目標日期=4/1(三)

  Scenario: 禮拜五斷點 - on_breakpoint
    Given ERP 資料:
      | 排程出貨日期 | 斷點    | ETA        | 日期算法 | 客戶需求地區 |
      | 2026-03-13   | 禮拜五 | 下下下週三 | ETA      | 429E         |
    When 計算目標日期
    Then 排程=3/13(五) → week_end=3/13(五)
    And on_breakpoint=True
    And 目標日期=4/8(三)（比正常多 7 天）

  # ── 同 cell 累加 ──

  Scenario: 多筆 ERP 填入同一 Commit cell
    Given 兩筆 ERP 對應同一料號和同一目標日期:
      | 客戶料號 | 淨需求 | 目標日期 |
      | 768290   | 150    | 4/9      |
      | 768290   | 10     | 4/9      |
    When 填入 Forecast
    Then 768290 在 4/9 欄位的值為 160,000（累加）

  # ── 安全防護 ──

  Scenario: 目標日期早於排程日期
    Given ERP 資料的計算結果 target_date < schedule_date
    When 計算目標日期
    Then 該筆 ERP 被跳過不填入

  Scenario: 淨需求為零或負數
    Given ERP 資料的淨需求=0
    When 處理 ERP 填入
    Then 該筆 ERP 被跳過不填入

  Scenario: 料號不存在於 Forecast
    Given ERP 的客戶料號在 Forecast 中找不到對應的 Commit row
    When 處理 ERP 填入
    Then 該筆 ERP 被跳過不填入

  # ── 已分配追蹤 ──

  Scenario: 防止重複填入
    Given ERP 某筆資料的「已分配」欄位為 "✓"
    When 處理 ERP 填入
    Then 該筆被跳過
    And 不重複填入 Forecast
```

---

## Feature 6: Transit 日期填入

```gherkin
Feature: Transit 在途資料填入 Forecast
  作為系統
  我需要將 Transit 的 ETA 日期和數量填入 Forecast 的 Commit row
  Transit 的填入優先於 ERP（先處理 Transit，再處理 ERP）

  Scenario: 正常 Transit 填入
    Given Transit 資料:
      | 客戶需求地區 | Ordered Item | Qty | ETA        |
      | 2680         | 768290       | 5   | 2026-04-09 |
    And Forecast 檔案 Plant=2680 有料號 768290 的 Commit row
    When 處理 Transit 填入
    Then 768290 的 Commit row 在 4/9 對應欄位填入 5,000

  Scenario: Transit ETA 落在 Weekly 範圍
    Given Transit ETA=2026-04-15
    And 4/15 不在 Daily 欄位中
    And 4/13~4/19 屬於某個 Weekly 欄位
    When 查找目標欄位
    Then 填入該 Weekly 欄位

  Scenario: Transit 處理順序
    Given 同時有 Transit 和 ERP 資料
    When 執行 Forecast 處理
    Then Transit 先填入
    And ERP 後填入
    And 同一 cell 的值為 Transit + ERP 的累加
```

---

## Feature 7: Forecast 多檔合併

```gherkin
Feature: 多個 Forecast 檔案合併處理
  作為光寶的 Forecast 管理人員
  當我上傳多個 Plant 的 Forecast 檔案時
  系統可以選擇合併成一個檔案統一處理

  Scenario: 23 個檔案合併為一個
    Given 上傳了 23 個 Forecast 檔案
    And 每個檔案有不同的 Plant code
    When 選擇「合併模式」處理
    Then 合併為一個 merged_forecast.xlsx
    And Row 1 為 headers（含 Plant 和 Buyer Code）
    And 所有資料行的 Column A = Plant code
    And 所有資料行的 Column B = Buyer Code
    And 原始欄位全部右移 2 位
    And 不同檔案相同日期的欄位被對齊

  Scenario: 合併模式下的 Plant Daily 限制
    Given 合併後 Daily 欄位範圍為所有 Plant 的聯集
    And Plant "429E" 原始 Daily 結束日期為 4/6
    And Plant "2680" 原始 Daily 結束日期為 4/8
    When 計算 429E 的 ERP 目標日期為 4/7
    Then 不匹配 Daily 欄位（因為 4/7 超過 429E 的 Daily 限制 4/6）
    And fallback 到 Weekly 或 GAP

  Scenario: 逐檔模式處理
    Given 上傳了 23 個 Forecast 檔案
    When 選擇「逐檔模式」處理
    Then 每個檔案獨立處理
    And 產出 23 個 forecast_{plant}_{buyer}.xlsx
    And 每個 Plant 只匹配自己的 ERP 資料

  Scenario: 單檔 fallback
    Given 只上傳 1 個 Forecast 檔案
    When 執行處理
    Then 使用單檔模式（merged_mode=False）
    And Plant code 從 C1 讀取
```

---

## Feature 8: 日期欄位查找策略

```gherkin
Feature: 目標日期對應 Forecast 欄位的查找策略
  作為系統
  我需要將計算出的目標日期對應到 Forecast 的正確欄位
  使用 Daily → Weekly → GAP Fallback → Monthly 的優先順序

  Background:
    Given Forecast 日期結構:
      | 類型    | 欄位範圍  | 日期範圍                  |
      | Daily   | K~AO      | 2026-03-09 ~ 2026-04-08  |
      | Weekly  | AP~BK     | 2026-04-13 ~ 2026-08-31  |
      | Monthly | BL~BQ     | 2026-09 ~ 2026-11        |

  Scenario: 精確匹配 Daily 欄位
    Given 目標日期=2026-03-19
    When 查找對應欄位
    Then 匹配到 3/19 的 Daily 欄位

  Scenario: 匹配 Weekly 範圍
    Given 目標日期=2026-04-15
    And 4/15 不在 Daily 欄位中
    And 4/13~4/19 屬於某個 Weekly 欄位
    When 查找對應欄位
    Then 匹配到 4/13 起始的 Weekly 欄位

  Scenario: GAP Fallback - 日期在 Daily-Weekly 空隙
    Given 目標日期=2026-04-09
    And Daily 最後日期=2026-04-08
    And Weekly 第一個日期=2026-04-13
    And 4/9 在空隙中（不屬於任何 Daily 或 Weekly）
    When 查找對應欄位
    Then Fallback 到第一個 Weekly 欄位 (4/13)

  Scenario: 03/07-start 組的較大 GAP
    Given Plant 429E 的 Daily 結束=2026-04-06
    And Weekly 開始=2026-04-13
    And 目標日期=2026-04-08（6天 GAP 區間）
    When 查找對應欄位
    Then GAP Fallback 到第一個 Weekly 欄位

  Scenario: Monthly Fallback
    Given 目標日期=2026-09-15
    And 9/15 不在 Daily 或 Weekly 中
    When 查找對應欄位
    Then 匹配到 9 月的 Monthly 欄位

  Scenario: 超出所有範圍
    Given 目標日期=2026-12-01
    And 12 月不在任何 Daily/Weekly/Monthly 中
    When 查找對應欄位
    Then 回傳 None
    And 該筆資料被跳過不填入
```

---

## 行為規格索引

| Feature | Scenario 數 | 涵蓋需求 |
|---------|-------------|---------|
| 客戶映射配置 | 12 | 前端 UI、CRUD、驗證、分頁、光寶專屬欄位 |
| Forecast 清理 | 3 | 上傳、清零、錯誤處理 |
| ERP 映射 | 4 | Type 11/32 查表、前綴解析、未匹配處理 |
| Transit 映射 | 2 | Type 11/32 在途映射 |
| ERP 日期填入 | 10 | 日期計算、on_breakpoint、累加、安全防護 |
| Transit 日期填入 | 3 | ETA 填入、Weekly 匹配、處理順序 |
| 多檔合併 | 4 | 合併/逐檔模式、Daily 限制、單檔 fallback |
| 日期欄位查找 | 5 | Daily/Weekly/GAP/Monthly 策略 |
| **合計** | **43** | |
