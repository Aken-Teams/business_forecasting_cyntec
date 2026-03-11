# FORECAST 數據處理系統 — BDD 行為驅動開發文件

**文件版本**: v1.0
**建立日期**: 2026-02-15
**機密等級**: 客戶文件

---

## Feature 1：使用者登入與權限

```gherkin
Feature: 使用者登入
  使用者需透過帳號密碼登入系統

  Scenario: 成功登入系統
    Given 使用者開啟系統登入頁面
    When 輸入正確的帳號與密碼
    And 點擊登入按鈕
    Then 系統驗證通過，進入主操作頁面
    And 系統記錄此次登入行為

  Scenario: 帳號或密碼錯誤
    Given 使用者開啟系統登入頁面
    When 輸入錯誤的帳號或密碼
    Then 系統顯示「帳號或密碼錯誤」提示
    And 無法進入系統

  Scenario: 登入逾時自動登出
    Given 使用者已登入系統超過 8 小時
    When 使用者嘗試操作任何功能
    Then 系統自動登出並導向登入頁面

  Scenario: 權限控制
    Given 使用者角色為「一般使用者」
    When 嘗試存取管理員功能（如使用者管理）
    Then 系統拒絕存取並導向首頁
```

---

## Feature 2：檔案上傳

```gherkin
Feature: ERP 淨需求上傳
  使用者上傳 ERP Excel 檔案，系統自動驗證格式

  Scenario: 上傳格式正確的 ERP 檔案
    Given 使用者已登入系統
    When 選擇符合模板格式的 ERP Excel 檔案並上傳
    Then 系統自動驗證檔案格式
    And 驗證通過，顯示上傳成功
    And 系統記錄此次上傳行為

  Scenario: 上傳格式不符的 ERP 檔案
    Given 使用者已登入系統
    When 選擇缺少必要欄位的 ERP Excel 檔案並上傳
    Then 系統格式驗證失敗
    And 顯示明確的錯誤提示（如「缺少欄位: 淨需求」）

Feature: Forecast 預測上傳
  支援多個 Forecast 檔案上傳與自動合併

  Scenario: 多檔上傳並合併
    Given 使用者已登入系統
    When 選擇 3 個 Forecast Excel 檔案上傳
    Then 系統驗證 3 個檔案格式皆通過
    And 使用者點擊合併
    Then 系統自動合併為一個檔案，保留原始格式

Feature: Transit 在途上傳

  Scenario: 需要 Transit 的客戶未上傳
    Given 使用者所屬客戶的部分廠區需要 Transit 檔案
    When 使用者跳過 Transit 上傳
    Then 系統提示「此廠區需要上傳在途清單」
```

---

## Feature 3：數據清理

```gherkin
Feature: Forecast 數據清理
  處理前自動清除舊有預測數據

  Scenario: 清除舊有供應數量
    Given 使用者已上傳 Forecast 檔案
    When 使用者點擊「數據清理」
    Then 系統自動清除供應數量相關欄位的舊數據
    And 保留 Excel 原有格式（字型、邊框、填色）
    And 顯示清理完成通知

  Scenario: 清除庫存數量
    Given Forecast 檔案中包含庫存數量資料
    When 執行數據清理
    Then 系統自動清除庫存數量相關數據
    And 其餘欄位不受影響
```

---

## Feature 4：Mapping 整合

```gherkin
Feature: 客戶 Mapping 設定與整合
  設定客戶與區域對應關係，並整合至 ERP / Transit 數據

  Scenario: 設定 Mapping 對應關係
    Given 使用者位於 Mapping 設定頁面
    When 編輯客戶對應設定（區域、排程斷點、ETD、ETA）
    And 點擊儲存
    Then 系統儲存設定，顯示成功訊息

  Scenario: 批次儲存多筆 Mapping
    Given 使用者編輯了 5 筆 Mapping 設定
    When 點擊儲存
    Then 系統一次儲存 5 筆，顯示「成功儲存 5 筆」

  Scenario: 整合 Mapping 至 ERP 數據
    Given 使用者已上傳 ERP 並完成數據清理
    And 已設定客戶 Mapping（區域、排程斷點、ETD、ETA）
    When 使用者點擊「Mapping 整合」
    Then 系統將 Mapping 資訊自動寫入 ERP 數據對應列
    And 顯示整合完成通知
```

---

## Feature 5：Forecast 預測處理

```gherkin
Feature: Forecast 預測運算
  將 ERP / Transit 數據比對填入 Forecast 報表

  Scenario: ERP 數據填入 Forecast
    Given 使用者已完成 ERP Mapping 整合
    And ERP 中有淨需求數據（客戶、料號、數量、出貨日）
    And Mapping 中有對應的排程斷點與 ETA 設定
    When 使用者點擊「Forecast 處理」
    Then 系統自動比對客戶與廠區
    And 依據排程斷點與 ETA 計算目標週別
    And 將淨需求數量填入 Forecast 對應的 ETA QTY 欄位
    And 顯示處理完成通知

  Scenario: 數量累加而非覆蓋
    Given Forecast 某位置已有預測數量 300
    And ERP 有另一筆同位置的淨需求 200
    When 執行 Forecast 處理
    Then 該位置數量變為 500（累加）

  Scenario: 已處理的數據不重複計算
    Given ERP 中某筆數據已在前次處理中分配完成
    When 再次執行 Forecast 處理
    Then 該筆數據被跳過，不會重複填入

  Scenario: Transit 在途數據一併處理
    Given 使用者已上傳 Transit 在途檔案
    When 執行 Forecast 處理
    Then 系統同時將 Transit 數據依 Mapping 計算目標週別
    And 填入 Forecast 對應的 ETA QTY 欄位
```

---

## Feature 6：結果下載

```gherkin
Feature: 處理結果下載
  使用者下載處理完成的 Excel 報表

  Scenario: 下載最終 Forecast 結果
    Given 使用者已完成所有處理階段
    When 點擊下載按鈕
    Then 系統提供以下檔案下載:
      | 檔案 | 說明 |
      | 清理後 Forecast | 經數據清理後的報表 |
      | 整合後 ERP | 已整合 Mapping 的 ERP 報表 |
      | 整合後 Transit | 已整合 Mapping 的 Transit 報表 |
      | Forecast 處理結果 | 最終預測報表 |
    And 系統記錄下載行為
```

---

## Feature 7：管理功能

```gherkin
Feature: 使用者管理
  管理員管理系統帳號

  Scenario: 新增使用者
    Given 管理員位於使用者管理頁面
    When 填入新使用者資訊（帳號、密碼、角色、公司）
    And 點擊新增
    Then 系統建立帳號，顯示成功訊息

  Scenario: 停用使用者
    Given 管理員位於使用者管理頁面
    When 將某使用者設為停用
    Then 該使用者無法再登入系統

Feature: IT 測試模式
  IT 人員可模擬客戶進行測試

  Scenario: 模擬客戶測試
    Given IT 人員位於 IT 儀表板
    When 選擇目標客戶進入測試模式
    Then 系統使用該客戶的專屬模板進行驗證
    And 測試結果不影響正式數據

Feature: 活動日誌查詢

  Scenario: 查詢操作紀錄
    Given IT 人員或管理員位於活動日誌頁面
    When 設定查詢條件（使用者、時間範圍、操作類型）
    And 點擊查詢
    Then 系統顯示符合條件的操作紀錄列表
```
