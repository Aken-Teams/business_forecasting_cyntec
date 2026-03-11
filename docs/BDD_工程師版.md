# FORECAST 數據處理系統 — BDD 行為驅動開發文件

**文件版本**: v1.0
**建立日期**: 2026-02-15
**機密等級**: 內部文件

---

## Feature 1：使用者認證

### Scenario 1.1：成功登入

```gherkin
Feature: 使用者登入
  使用者需透過帳號密碼登入系統才能操作

  Scenario: 使用正確帳密登入
    Given 使用者位於 "/login" 頁面
    And 資料庫中存在帳號 "user01"，密碼雜湊為 SHA-256(salt + "password123")
    When 使用者輸入帳號 "user01" 與密碼 "password123"
    And 點擊登入按鈕，POST 至 "/api/login"
    Then 系統回傳 {"success": true, "user": {...}}
    And Session 儲存 user_id, username, role, company, login_time
    And 頁面導向 "/" 主操作頁
    And activity_logs 新增一筆 action="login" 記錄（含 IP 與 User-Agent）
```

### Scenario 1.2：登入失敗

```gherkin
  Scenario: 使用錯誤密碼登入
    Given 使用者位於 "/login" 頁面
    When 使用者輸入帳號 "user01" 與密碼 "wrongpass"
    And 點擊登入按鈕，POST 至 "/api/login"
    Then 系統回傳 {"success": false, "message": "帳號或密碼錯誤"}
    And Session 不建立
    And activity_logs 新增一筆 action="login_failed" 記錄
```

### Scenario 1.3：Session 逾時

```gherkin
  Scenario: Session 超過 8 小時自動登出
    Given 使用者已登入且 Session login_time 為 8 小時前
    When 使用者嘗試存取任何 @login_required 路由
    Then 系統清除 Session
    And 回傳 302 導向 "/login"
```

### Scenario 1.4：Pegatron 限制

```gherkin
  Scenario: 非 Pegatron 使用者無法登入
    Given 資料庫中使用者 "quanta_user" 之 company 為 "Quanta"
    When 使用者輸入正確帳密
    Then 系統回傳 {"success": false, "message": "僅限 Pegatron 使用者登入"}
```

---

## Feature 2：檔案上傳

### Scenario 2.1：ERP 上傳成功

```gherkin
Feature: ERP 淨需求上傳
  使用者上傳 ERP Excel 檔案，系統進行格式驗證

  Scenario: 上傳符合 Pegatron 模板的 ERP 檔案
    Given 使用者已登入，company="Pegatron"
    And compare/pegatron/erp.xlsx 模板存在
    When 使用者上傳 "erp_data.xlsx"，POST multipart/form-data 至 "/upload_erp"
    Then 系統讀取 compare/pegatron/erp.xlsx 模板
    And 比對上傳檔案欄位結構（客戶簡稱、客戶PO、客料、淨需求、送貨地點、排程出貨日、倉庫）
    And 驗證通過
    And 檔案存入 uploads/{user_id}/{timestamp}/erp.xlsx
    And upload_records 新增一筆 file_type="erp", validation_ok=true
    And activity_logs 新增一筆 action="upload_erp"
    And 回傳 {"success": true, "filename": "erp.xlsx"}
```

### Scenario 2.2：ERP 格式驗證失敗

```gherkin
  Scenario: 上傳缺少必要欄位的 ERP 檔案
    Given 使用者已登入
    When 使用者上傳缺少「淨需求」欄位的 Excel 檔案至 "/upload_erp"
    Then 系統格式驗證失敗
    And upload_records 新增一筆 validation_ok=false, error_message="缺少欄位: 淨需求"
    And activity_logs 新增一筆 action="upload_erp_failed"
    And 回傳 {"success": false, "error": "格式驗證失敗: 缺少欄位: 淨需求"}
```

### Scenario 2.3：Forecast 多檔上傳與合併

```gherkin
Feature: Forecast 多檔上傳
  支援多個 Forecast 檔案上傳後自動合併

  Scenario: 上傳 3 個 Forecast 檔案並合併
    Given 使用者已登入
    When 使用者上傳 3 個 Forecast .xlsx 檔案至 "/upload_forecast"
    Then 系統依序驗證 3 個檔案格式
    And 3 個檔案皆通過驗證
    And 使用者觸發 POST "/merge_forecast"
    Then excel_processor.merge_excel_files() 合併 3 個檔案
    And 合併時保留 merged cells 格式
    And 輸出 merged_forecast.xlsx 至 uploads/{user_id}/{timestamp}/
    And 回傳 {"success": true, "merged_file": "merged_forecast.xlsx"}
```

### Scenario 2.4：Transit 上傳（Pegatron 必要）

```gherkin
Feature: Transit 在途上傳

  Scenario: Pegatron 使用者需上傳 Transit
    Given 使用者已登入，company="Pegatron"
    And customer_mappings 中存在 requires_transit=true 的廠區
    When 使用者跳過 Transit 上傳直接進入處理階段
    Then 系統提示 "此廠區需要上傳在途清單"
```

---

## Feature 3：數據清理

### Scenario 3.1：清理供應數量

```gherkin
Feature: Forecast 數據清理
  清除舊有供應數量與庫存數據

  Scenario: K 欄為「供應數量」時清除 L~AW 欄
    Given 已上傳 Forecast 檔案
    And Forecast 中第 10 列 K 欄值為 "供應數量"
    When 使用者點擊「數據清理」，POST 至 "/process/cleanup"
    Then 系統以 openpyxl 開啟 Forecast
    And 第 10 列 L 欄至 AW 欄所有儲存格值清空
    And 儲存格格式（字型、邊框、填色）保留不變
    And 存檔為 processed/{user_id}/{timestamp}/cleaned_forecast.xlsx
    And process_records 新增 process_type="cleanup", status="success"
```

### Scenario 3.2：清理庫存數量

```gherkin
  Scenario: I 欄含「庫存數量」時清除下一列 I 欄
    Given Forecast 中第 15 列 I 欄值含 "庫存數量"
    When 執行數據清理
    Then 第 16 列 I 欄值被清空
    And 其餘欄位不受影響
```

---

## Feature 4：Mapping 整合

### Scenario 4.1：讀取 Mapping 並整合至 ERP

```gherkin
Feature: ERP Mapping 整合
  將客戶 Mapping 設定整合至 ERP 數據

  Scenario: 整合 Mapping 欄位至 ERP
    Given 使用者已上傳 ERP 並完成清理
    And customer_mappings 中有以下設定:
      | customer_name | region        | schedule_breakpoint | etd    | eta      |
      | 和碩          | MAINTEK-新寧   | 禮拜三              | 下週二  | 下下週二  |
    When 使用者點擊「Mapping 整合」，POST 至 "/process/erp_mapping"
    Then 系統從 DB 讀取 get_customer_mappings(user_id)
    And ERP 中客戶簡稱="和碩" 且送貨地點="MAINTEK-新寧" 的列
    And 寫入 region, schedule_breakpoint, etd, eta 欄位
    And 存檔為 integrated_erp.xlsx
```

### Scenario 4.2：Mapping 設定 CRUD

```gherkin
Feature: Mapping 管理

  Scenario: 批次儲存 Mapping
    Given 使用者位於 "/mapping" 頁面
    When 使用者編輯 5 筆 Mapping 後點擊儲存
    And 前端 POST JSON 至 "/api/mapping/list"
    Then database.save_customer_mappings() 儲存 5 筆
    And UNIQUE KEY (user_id, customer_name, region) 衝突時執行 UPDATE
    And 回傳 {"success": true, "count": 5}
```

---

## Feature 5：Forecast 預測處理

### Scenario 5.1：ERP 數據填入 Forecast ETA QTY

```gherkin
Feature: Forecast 預測運算
  將 ERP/Transit 數據比對填入 Forecast 的 ETA QTY 列

  Scenario: Pegatron ERP 數據比對並填入
    Given 已完成 ERP Mapping 整合
    And ERP 中有一筆:
      | 客戶簡稱 | Line客戶採購單號 | 客料           | 淨需求 | 排程出貨日  | 送貨地點      |
      | 和碩     | 3A32-A00Y       | 0703-00AG000  | 500    | 2026/01/20 | MAINTEK-新寧  |
    And Mapping: schedule_breakpoint="禮拜三", eta="下下週二"
    And Forecast 中存在 Plant=3A32, MRP_ID=A00Y, PN_Model=0703-00AG000 的 Block
    When 使用者點擊「Forecast 處理」，POST 至 "/run_forecast"
    Then pegatron_forecast_processor 計算:
      | 步驟 | 計算 |
      | 排程出貨日 2026/01/20 所在週 | 週一=01/19 ~ 週日=01/25 |
      | 排程斷點=禮拜三 | 週末日=01/22 (三) |
      | ETA=下下週二 | 01/22 + 14天 → 02/03 週 → 02/03 (二) |
    And 找到 Forecast 中日期 2026/02/02 (該週一) 所在的週欄位
    And 該 Block 的 ETA QTY 列對應欄位累加 500
    And ERP 該筆標記為「已分配」(✓)
    And 存檔為 forecast_result.xlsx
```

### Scenario 5.2：數量累加（非覆蓋）

```gherkin
  Scenario: 同一目標儲存格有多筆數據時累加
    Given Forecast 某 Block 之 ETA QTY 列，2026/02/02 週欄位已有值 300
    And ERP 新一筆同 Block 同週 淨需求=200
    When 執行 Forecast 處理
    Then 該儲存格值變為 300 + 200 = 500（累加而非覆蓋）
```

### Scenario 5.3：已分配記錄跳過

```gherkin
  Scenario: 已分配的 ERP 記錄不重複處理
    Given ERP 中某筆記錄已標記「已分配」(✓)
    When 再次執行 Forecast 處理
    Then 該筆記錄被跳過，不重複填入數量
```

### Scenario 5.4：Transit 數據整合

```gherkin
  Scenario: Transit 在途數據填入 Forecast
    Given 已上傳 Transit 檔案
    And Transit 中有一筆:
      | Line客戶採購單號 | Ordered Item   | Qty | 送貨地點      |
      | 3A32-A00Y       | 0703-00AG000  | 200 | MAINTEK-新寧  |
    And Transit KEY = Line客戶採購單號 + Ordered Item
    When 執行 Forecast 處理
    Then 系統將 Transit 數據同樣依 Mapping 計算目標週
    And 填入對應 Block 的 ETA QTY 列
```

---

## Feature 6：結果下載

### Scenario 6.1：下載處理結果

```gherkin
Feature: 結果下載

  Scenario: 下載 Forecast 處理結果
    Given 使用者已完成所有處理階段
    And processed/{user_id}/{timestamp}/ 下存在 forecast_result.xlsx
    When 使用者點擊下載按鈕，GET "/download/forecast_result.xlsx"
    Then 系統回傳檔案，Content-Type 為 application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
    And activity_logs 新增 action="download_file"
```

---

## Feature 7：管理功能

### Scenario 7.1：使用者管理

```gherkin
Feature: 使用者管理

  Scenario: Admin 新增使用者
    Given 使用者角色為 "admin"，位於 "/users_manage"
    When POST "/api/users" body={"username":"new_user","password":"pass123","role":"user","company":"Pegatron"}
    Then database.create_user() 建立使用者
    And password_hash = SHA-256(PASSWORD_SALT + "pass123")
    And 回傳 {"success": true, "user_id": 10}

  Scenario: 一般使用者無法存取管理頁面
    Given 使用者角色為 "user"
    When GET "/users_manage"
    Then @admin_required 裝飾器攔截
    And 回傳 403 或導向首頁
```

### Scenario 7.2：IT 測試模式

```gherkin
Feature: IT 測試模式

  Scenario: IT 人員模擬客戶測試
    Given 使用者角色為 "it"，位於 "/it" 儀表板
    When 選擇目標客戶 "Pegatron" 進入測試模式
    Then 系統使用 compare/pegatron/ 下的模板進行格式驗證
    And 測試結果存入 IT 使用者自己的 uploads/{it_user_id}/ 目錄
```

### Scenario 7.3：活動日誌查詢

```gherkin
Feature: 活動日誌

  Scenario: 查詢特定使用者的上傳記錄
    Given 使用者角色為 "it" 或 "admin"
    When GET "/api/logs?user_id=5&action=upload_erp&start=2026-01-01&end=2026-01-31"
    Then 系統查詢 activity_logs WHERE user_id=5 AND action='upload_erp' AND created_at BETWEEN ...
    And 回傳 {"success": true, "logs": [...], "total": 15}
```

---

## Feature 8：瀏覽器相容性

### Scenario 8.1：Chrome 阻擋

```gherkin
Feature: 瀏覽器偵測

  Scenario: Chrome 使用者被阻擋
    Given 使用者以 Chrome 瀏覽器開啟系統
    When browser-check.js 偵測 User-Agent 含 "Chrome" 且不含 "Edg"
    Then 顯示不可關閉的 Modal
    And Modal 提供「以 Edge 開啟」按鈕
    And 使用者無法操作背景頁面
```
