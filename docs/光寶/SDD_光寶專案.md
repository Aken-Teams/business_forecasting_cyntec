# FORECAST 數據處理系統 — 光寶科技客製化擴展 SDD

**文件版本**: v1.0
**建立日期**: 2026-03-11
**專案名稱**: FORECAST 數據處理系統 — 光寶科技 (Liteon) 客製化擴展
**機密等級**: 客戶文件

---

## 1. 設計概述

### 1.1 擴展範圍

本次客製化擴展在現有系統架構上新增光寶科技專屬處理邏輯，涵蓋以下模組：

| 擴展模組 | 說明 |
|----------|------|
| ERP Mapping 整合 | 新增雙訂單類型 (11/32) 之差異化比對邏輯 |
| Transit Mapping 整合 | 新增在途數據反向查詢 ERP 送貨地點/倉庫之比對邏輯 |
| Forecast 處理引擎 | 新增專屬處理器，支援 Daily+Weekly+Monthly 三段式日期結構 |
| Mapping 設定介面 | 擴展四個新欄位（訂單型態、送貨地點、倉庫、日期算法） |
| 下載模組 | 支援多檔案批量下載 |

### 1.2 設計原則

| 原則 | 說明 |
|------|------|
| 零影響擴展 | 新增功能不修改現有客戶之任何處理邏輯 |
| 帳號自動識別 | 系統依登入帳號自動切換對應之處理邏輯 |
| 模組化設計 | 光寶專屬處理器為獨立模組，易於維護與測試 |
| 格式保留 | 處理過程完整保留 Excel 原始格式 |

---

## 2. 系統架構

### 2.1 模組關係圖

```mermaid
graph TB
    subgraph 前端介面
        U1["檔案上傳（多檔）"]
        U2["Mapping 設定介面"]
        U3["結果下載（批量）"]
    end

    subgraph 後端服務
        ROUTER["客戶識別與路由分派"]
        LITEON["光寶專屬處理邏輯"]
        OTHER["其他客戶處理邏輯"]

        subgraph ENGINE["光寶 Forecast 處理引擎"]
            TR["Transit 填入模組"]
            ER["ERP 填入模組"]
        end
    end

    U1 --> ROUTER
    U2 --> ROUTER
    U3 --> ROUTER
    ROUTER --> LITEON
    ROUTER --> OTHER
    LITEON --> ENGINE
```

### 2.2 處理流程架構

```mermaid
flowchart TD
    ERP["ERP 淨需求"] --> MAP_E["ERP Mapping 整合"]
    FC["Forecast 多檔預測"] --> CLEAN["數據清理"]
    TR["Transit 在途"] --> MAP_T["Transit Mapping 整合"]

    MAP_E --> ENGINE
    CLEAN --> ENGINE
    MAP_T --> ENGINE

    subgraph ENGINE["Forecast 處理引擎"]
        direction TB
        S1["① 載入所有數據"]
        S2["② Transit 數據填入"]
        S3["③ ERP 數據填入"]
        S4["④ 分配狀態回寫"]
        S5["⑤ 結果輸出"]
        S1 --> S2 --> S3 --> S4 --> S5
    end

    ENGINE --> OUT["多檔結果 Excel 輸出"]
```

---

## 3. Mapping 整合設計

### 3.1 ERP Mapping 流程

```mermaid
flowchart TD
    A["ERP 每一筆資料"] --> B["讀取客戶簡稱"]
    B --> C{"在 Mapping 表中<br/>搜尋匹配？"}
    C -->|找到匹配| D{"判斷訂單型態"}
    C -->|未找到| SKIP["跳過"]

    D -->|"11（一般訂單）"| E{"比對 ERP 送貨地點<br/>= Mapping 送貨地點？"}
    D -->|"32（HUB 調撥）"| F{"比對 ERP 倉庫<br/>= Mapping 倉庫代碼？"}

    E -->|匹配| G["填入廠區代碼、斷點、ETD、ETA"]
    E -->|不匹配| SKIP
    F -->|匹配| G
    F -->|不匹配| SKIP

    G --> NEXT["下一筆"]
    SKIP --> NEXT
```

### 3.2 Transit Mapping 流程

```mermaid
flowchart TD
    A["Transit 每一筆資料"] --> B["讀取 Transit 地點 (Location)"]
    B --> C{"在 ERP 中反查<br/>此地點對應的訂單型態"}

    C -->|"匹配 ERP 送貨地點<br/>→ 訂單型態 11"| D["以 (客戶簡稱 + 送貨地點)<br/>查詢 Mapping → 取得廠區代碼"]
    C -->|"匹配 ERP 倉庫<br/>→ 訂單型態 32"| E["以 (客戶簡稱 + 倉庫)<br/>查詢 Mapping → 取得廠區代碼"]
    C -->|均不匹配| SKIP["跳過"]

    D --> F["填入客戶需求地區（廠區代碼）"]
    E --> F
    F --> NEXT["下一筆"]
    SKIP --> NEXT
```

---

## 4. Forecast 處理引擎設計

### 4.1 處理總流程

```mermaid
flowchart TD
    IN["輸入：Forecast × N + 已整合 ERP + 已整合 Transit"]
    IN --> LOOP["對每個 Forecast 檔案"]
    LOOP --> S1["① 載入 Forecast<br/>識別廠區代碼 (Plant)、採購員代碼 (Buyer)"]
    S1 --> S2["② 解析日期標頭<br/>建立日期對照表（Daily / Weekly / Monthly）"]
    S2 --> S3["③ 建立料號索引<br/>每個料號對應 Commit 列位置"]
    S3 --> S4["④ Transit 填入<br/>比對廠區+料號，填入 Commit 列"]
    S4 --> S5["⑤ ERP 填入<br/>比對廠區+料號+日期計算，填入 Commit 列"]
    S5 --> S6["⑥ 回寫分配狀態<br/>標記已使用之 ERP/Transit 資料"]
    S6 --> S7["⑦ 儲存結果檔案<br/>forecast_{廠區}_{採購員}.xlsx"]
    S7 --> LOOP
```

### 4.2 Forecast 檔案結構解析

```mermaid
block-beta
    columns 3
    block:HEADER:3
        columns 2
        C1["C1 = 廠區代碼 (Plant)"]
        E1["E1 = 採購員代碼 (Buyer)"]
    end
    block:DATES:3
        columns 3
        DAILY["K ~ AO<br/>每日日期（31 天）"]
        WEEKLY["AP ~ BK<br/>每週日期（22 週）"]
        MONTHLY["BL ~ BQ<br/>每月日期（6 個月）"]
    end
    block:DATA:3
        columns 3
        MAT["B 欄<br/>料號 (Material)"]
        MEASURE["C 欄<br/>資料類型"]
        VALUES["K~BQ 欄<br/>數量"]
    end
    block:ROWS:3
        columns 3
        R1["Demand（唯讀）"]
        R2["Commit（填入目標）"]
        R3["Accumulate Shortage（唯讀）"]
    end

    style R2 fill:#e8f5e9,stroke:#2e7d32,stroke-width:2px
```

> **Sheet**: `Daily+Weekly+Monthly` | **Row 7** = 日期標頭 | **Row 8+** = 數據列（每個料號 3 列一組）

### 4.3 日期對應策略

系統將目標日期對應至 Forecast 欄位時，採用**精度遞減**策略：

```mermaid
flowchart TD
    START["目標日期：2026-03-19"] --> D{"① 搜尋 Daily 區段<br/>(K~AO)"}
    D -->|"找到 3/19"| D_OK["填入該欄 ✓"]
    D -->|"無對應"| W{"② 搜尋 Weekly 區段<br/>(AP~BK)"}
    W -->|"找到 3/19 所屬週"| W_OK["填入該週欄 ✓"]
    W -->|"無對應"| M{"③ 搜尋 Monthly 區段<br/>(BL~BQ)"}
    M -->|"找到 3 月"| M_OK["填入該月欄 ✓"]
    M -->|"無對應"| SKIP["跳過此筆"]
```

### 4.4 ERP 日期計算邏輯

ERP 目標日期的計算基於排程出貨日期、斷點、及 ETD/ETA 文字描述：

```mermaid
flowchart TD
    IN["輸入：排程出貨日期 = 3/10（週二）<br/>排程斷點 = 禮拜一<br/>日期算法 = ETA<br/>ETA 文字 = 下週四"]

    IN --> S1["① 從排程出貨日期往後<br/>找到下一個「禮拜一」<br/>→ 3/16（週一）= 本週期終點"]
    S1 --> S2["② 從週期終點解析「下週四」<br/>→ 下一個週期的週四 = 3/19"]
    S2 --> S3{"③ 驗證：3/19 ≥ 3/10？"}
    S3 -->|"✓ 通過"| RESULT["結果：目標日期 = 2026-03-19"]
    S3 -->|"✗ 不通過"| SKIP["跳過此筆（不填入）"]
```

**支援的日期文字格式**：

| 文字 | 說明 | 範例 |
|------|------|------|
| 本週X | 當前週期內的星期X | 本週五 |
| 下週X | 下一個週期的星期X | 下週四 |
| 下下週X | 再下一個週期的星期X | 下下週二 |

**安全機制**：計算結果若早於排程出貨日期，該筆數據將被跳過（不填入）。

### 4.5 數量填入規則

| 規則 | 說明 |
|------|------|
| 單位轉換 | ERP/Transit 原始數量 × 1000 後填入 Forecast |
| 數量累加 | 同一儲存格若有多筆來源，數量自動累加 |
| 保留原值 | 若 Commit 儲存格已有原始數值，新增數量疊加於上 |
| 零值處理 | 數量為 0 之資料不填入 |

### 4.6 分配追蹤機制

```mermaid
sequenceDiagram
    participant F1 as Forecast 1 (15K0)
    participant ERP as ERP/Transit 數據
    participant F2 as Forecast 2 (15K0)
    participant F3 as Forecast 3 (F820)

    F1->>ERP: 使用 ERP #1、#3、#7
    ERP-->>F1: 填入 Commit
    F1->>ERP: 標記「已分配 ✓」
    F1->>ERP: 使用 Transit #2、#5
    ERP-->>F1: 填入 Commit
    F1->>ERP: 標記「已分配 ✓」

    F2->>ERP: 讀取 ERP #1、#3、#7
    ERP-->>F2: 已分配 → 跳過
    F2->>ERP: 使用 ERP #12、#15
    ERP-->>F2: 填入 Commit
    F2->>ERP: 標記「已分配 ✓」

    F3->>ERP: 不同廠區(F820)，匹配不同筆數
    ERP-->>F3: 填入 Commit

    Note over F1,F3: 每筆 ERP/Transit 最多只被使用一次<br/>分配狀態回寫至來源檔案供稽核參考
```

---

## 5. Mapping 設定介面設計

### 5.1 欄位配置

光寶科技之 Mapping 表格新增四個專屬欄位，介面自動判斷並展開：

| 客戶簡稱 | 訂單型態 | 送貨地點 | 倉庫 | 廠區 | 排程斷點 | ETD | ETA | 日期算法 | Transit需求 |
|---------|:-------:|---------|------|------|---------|------|------|:-------:|:----------:|
| 光寶科技... | 11 | TB01 | | 15K0 | 禮拜一 | 下週四 | 下週四 | ETA | 是 |
| 光寶科技... | 32 | | HUB_A | 15K0 | 禮拜一 | 下週四 | 下週四 | ETA | 是 |

### 5.2 操作方式

| 操作 | 說明 |
|------|------|
| 新增 | 點擊「新增客戶」按鈕，新增一列空白設定 |
| 編輯 | 直接於表格欄位內編輯 |
| 儲存 | 支援單筆及批次儲存 |
| 刪除 | 刪除不需要的設定列 |
| 分頁 | 每頁顯示 10 筆，支援分頁切換 |

---

## 6. 輸出檔案設計

### 6.1 結果檔案命名規則

命名格式：`forecast_{Plant}_{Buyer}.xlsx`

| 範例檔名 | 廠區 | 採購員 |
|----------|------|--------|
| forecast_15K0_P43.xlsx | 15K0 | P43 |
| forecast_15K0_P49.xlsx | 15K0 | P49 |
| forecast_F820_T12.xlsx | F820 | T12 |

### 6.2 結果檔案內容

| 項目 | 說明 |
|------|------|
| 格式保留 | 完整保留原始 Excel 格式（字型、色彩、框線、合併儲存格） |
| Demand 列 | 保持原始數據不變 |
| Commit 列 | 填入 ERP 及 Transit 之供應承諾數量 |
| Shortage 列 | 保持原始數據不變 |

### 6.3 批量下載

處理完成後，下載區域顯示所有結果檔案清單：

| 檔案名稱 | ERP 填入 | Transit 填入 | 操作 |
|----------|:-------:|:-----------:|:----:|
| **批量下載全部檔案（共 23 個）** | | | **下載全部** |
| forecast_15K0_P43.xlsx | 52 筆 | 3 筆 | 下載 |
| forecast_15K0_P49.xlsx | 48 筆 | 2 筆 | 下載 |
| forecast_F820_T12.xlsx | 61 筆 | 5 筆 | 下載 |
| ... | | | |

---

## 7. 數據流總覽

```mermaid
flowchart LR
    subgraph 上傳階段
        E1["ERP .xlsx<br/>格式驗證 / 欄位檢查"]
        F1["Forecast .xlsx × N<br/>格式驗證 / 數據清理"]
        T1["Transit .xlsx<br/>格式驗證 / 欄位檢查"]
    end

    subgraph 處理階段
        M["Mapping 整合<br/>ERP整合 + Transit整合"]
        P["Forecast 處理<br/>Transit填入 / ERP填入 / 分配追蹤"]
        M --> P
    end

    subgraph 輸出階段
        O1["forecast_15K0_P43.xlsx"]
        O2["forecast_15K0_P49.xlsx"]
        O3["forecast_F820_T12.xlsx"]
        O4["integrated_erp/transit.xlsx"]
    end

    E1 --> M
    F1 --> M
    T1 --> M
    P --> O1
    P --> O2
    P --> O3
    P --> O4
```

---

## 8. 品質保證

### 8.1 數據正確性保證

| 保證項目 | 機制 |
|----------|------|
| 廠區比對一致性 | ERP/Transit 之廠區代碼必須完全匹配 Forecast 之 Plant |
| 料號比對一致性 | ERP 客戶料號 / Transit 訂單品項必須完全匹配 Forecast 料號 |
| 日期安全檢查 | 計算日期不得早於排程出貨日期，違反則跳過 |
| 防重複填入 | 分配追蹤確保同一筆數據只被使用一次 |
| 數量單位統一 | 統一以 × 1000 轉換，確保單位一致 |

### 8.2 異常處理

| 異常情境 | 處理方式 |
|----------|---------|
| ERP/Transit 找不到匹配的 Forecast 料號 | 跳過該筆，不影響其他筆 |
| 日期欄位缺失或格式錯誤 | 跳過該筆，記錄至處理統計 |
| 目標日期超出 Forecast 日期範圍 | 嘗試週/月遞減匹配，仍無匹配則跳過 |
| 單一檔案處理失敗 | 記錄失敗原因，繼續處理其餘檔案 |
| Mapping 設定不完整 | 該筆數據跳過，不影響其他已設定之筆 |

---

## 9. 與現有系統整合

### 9.1 共用元件

| 元件 | 說明 |
|------|------|
| 使用者認證 | 使用系統統一認證機制，光寶為獨立帳號 |
| 檔案上傳模組 | 共用既有上傳與格式驗證架構 |
| Mapping 資料庫 | 共用 Mapping 資料表結構，擴展新欄位 |
| 活動日誌 | 所有操作自動記錄至統一日誌系統 |
| 檔案管理 | 共用自動清理機制 |
| IT 測試模式 | IT 人員可透過測試模式模擬光寶帳號進行測試 |

### 9.2 獨立元件

| 元件 | 說明 |
|------|------|
| Forecast 處理引擎 | 光寶專屬，獨立模組 |
| ERP Mapping 邏輯 | 雙訂單類型比對，為光寶特有 |
| Transit Mapping 邏輯 | 反向查詢 ERP 地點，為光寶特有 |
| Mapping 介面擴展 | 四個新欄位僅光寶帳號顯示 |

---

## 10. 術語表

| 術語 | 說明 |
|------|------|
| Plant（廠區） | Forecast 中的廠區代碼，用於區分不同生產據點 |
| Buyer（採購員） | Forecast 中的採購員代碼 |
| Region | Mapping 中的廠區代碼，對應 Forecast 的 Plant |
| 訂單型態 11 | 一般訂單，以送貨地點為比對鍵 |
| 訂單型態 32 | HUB 調撥訂單，以倉庫代碼為比對鍵 |
| 排程斷點 | 定義每個排程週期的結束日（星期幾） |
| ETD / ETA | 預計出發日 / 預計到達日 |
| Commit | Forecast 中記錄供應承諾量的資料列 |
| Daily / Weekly / Monthly | Forecast 日期結構的三個精度層級 |
| 已分配 | 用於標記已被處理過的 ERP/Transit 資料 |
| B/S 架構 | 瀏覽器/伺服器架構 (Browser/Server) |
