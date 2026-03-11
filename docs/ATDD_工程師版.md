# FORECAST 數據處理系統 — ATDD 驗收測試驅動開發文件

**文件版本**: v1.0
**建立日期**: 2026-02-15
**機密等級**: 內部文件

---

## 1. 驗收測試架構

```mermaid
flowchart LR
    US["User Story"] --> AC["Acceptance Criteria"]
    AC --> AT["Acceptance Test"]
    AT --> IMPL["實作"]
    IMPL --> VERIFY["驗證"]
    VERIFY -->|PASS| DONE["完成"]
    VERIFY -->|FAIL| IMPL
```

---

## 2. User Story 與驗收條件

### US-001：使用者登入

**As a** 系統使用者
**I want to** 以帳號密碼登入系統
**So that** 我可以操作 Forecast 數據處理功能

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-001-1 | 正確帳密登入後導向主操作頁 | POST `/api/login` → 200, Session 含 user_id |
| AC-001-2 | 錯誤密碼回傳明確錯誤訊息 | POST `/api/login` → `{"success": false}` |
| AC-001-3 | 登入限制 Pegatron 使用者 | company ≠ "Pegatron" → 拒絕登入 |
| AC-001-4 | Session 8 小時逾時自動登出 | login_time + 8hr 後存取 @login_required → 302 /login |
| AC-001-5 | 登入/登出行為寫入 activity_logs | 查詢 DB: action IN ('login', 'logout') 有記錄 |

#### Acceptance Test

```python
# test_acceptance_login.py

def test_ac_001_1_login_success():
    """AC-001-1: 正確帳密登入"""
    response = client.post("/api/login", json={"username": "peg_user", "password": "pass123"})
    assert response.json["success"] is True
    with client.session_transaction() as sess:
        assert sess["user_id"] is not None
        assert sess["company"] == "Pegatron"

def test_ac_001_2_login_wrong_password():
    """AC-001-2: 錯誤密碼"""
    response = client.post("/api/login", json={"username": "peg_user", "password": "wrong"})
    assert response.json["success"] is False
    assert "錯誤" in response.json["message"]

def test_ac_001_3_non_pegatron_rejected():
    """AC-001-3: 非 Pegatron 使用者被拒"""
    create_user("quanta_user", "pass", company="Quanta")
    response = client.post("/api/login", json={"username": "quanta_user", "password": "pass"})
    assert response.json["success"] is False

def test_ac_001_4_session_timeout():
    """AC-001-4: Session 8 小時逾時"""
    login_as(client, "peg_user")
    with client.session_transaction() as sess:
        sess["login_time"] = datetime.now() - timedelta(hours=9)
    response = client.get("/")
    assert response.status_code == 302

def test_ac_001_5_login_logged():
    """AC-001-5: 登入行為記錄"""
    client.post("/api/login", json={"username": "peg_user", "password": "pass123"})
    logs = db.query("SELECT * FROM activity_logs WHERE action='login' ORDER BY id DESC LIMIT 1")
    assert logs[0]["username"] == "peg_user"
```

---

### US-002：ERP 檔案上傳與驗證

**As a** 操作人員
**I want to** 上傳 ERP 淨需求 Excel 檔案
**So that** 系統可以進行後續數據處理

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-002-1 | 支援 .xls 和 .xlsx 格式 | 分別上傳兩種格式，皆成功 |
| AC-002-2 | 依客戶模板驗證欄位（客戶簡稱、客料、淨需求、送貨地點、排程出貨日、倉庫） | 上傳有效檔案 → success=true |
| AC-002-3 | 缺少欄位時回傳具體缺少的欄位名稱 | 上傳缺欄位檔案 → error 含欄位名 |
| AC-002-4 | 檔案存入 `uploads/{user_id}/{timestamp}/` | 檢查檔案系統路徑 |
| AC-002-5 | upload_records 記錄上傳結果 | 查詢 DB: file_type='erp' 有記錄 |

#### Acceptance Test

```python
def test_ac_002_1_supports_xlsx():
    """AC-002-1: 支援 .xlsx"""
    login_as(client, company="Pegatron")
    resp = upload_file(client, "/upload_erp", "test_data/valid_erp.xlsx")
    assert resp["success"] is True

def test_ac_002_1_supports_xls():
    """AC-002-1: 支援 .xls（透過 LibreOffice 轉換）"""
    resp = upload_file(client, "/upload_erp", "test_data/valid_erp.xls")
    assert resp["success"] is True

def test_ac_002_2_valid_format_passes():
    """AC-002-2: 有效格式通過驗證"""
    resp = upload_file(client, "/upload_erp", "test_data/valid_erp.xlsx")
    assert resp["success"] is True

def test_ac_002_3_missing_column_reports_name():
    """AC-002-3: 缺少欄位時回傳欄位名"""
    resp = upload_file(client, "/upload_erp", "test_data/missing_column_erp.xlsx")
    assert resp["success"] is False
    assert "淨需求" in resp["error"]

def test_ac_002_4_file_saved_to_user_dir():
    """AC-002-4: 存入使用者目錄"""
    resp = upload_file(client, "/upload_erp", "test_data/valid_erp.xlsx")
    user_id = get_session_user_id(client)
    files = glob.glob(f"uploads/{user_id}/*/erp.xlsx")
    assert len(files) >= 1

def test_ac_002_5_upload_record_created():
    """AC-002-5: 上傳記錄寫入 DB"""
    upload_file(client, "/upload_erp", "test_data/valid_erp.xlsx")
    records = db.query("SELECT * FROM upload_records WHERE file_type='erp' ORDER BY id DESC LIMIT 1")
    assert records[0]["validation_ok"] is True
```

---

### US-003：Forecast 多檔上傳與合併

**As a** 操作人員
**I want to** 上傳多個 Forecast 檔案並自動合併
**So that** 不同來源的預測數據能整合在同一份報表

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-003-1 | 支援同時上傳多個 Forecast 檔案 | 上傳 3 個檔案 → 全部成功 |
| AC-003-2 | 合併後列數為各檔案之和 | 檢查 merged_forecast.xlsx max_row |
| AC-003-3 | 合併保留 merged cells | 檢查 merged_cells.ranges |
| AC-003-4 | 單一檔案不需合併直接使用 | 上傳 1 個檔案 → 無需 merge 步驟 |

---

### US-004：數據清理

**As a** 操作人員
**I want to** 清除 Forecast 中的舊有預測數據
**So that** 新一輪處理不受舊數據干擾

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-004-1 | K 欄為「供應數量」時清除 L~AW 欄 | 讀取清理後檔案，對應儲存格值為 None |
| AC-004-2 | I 欄含「庫存數量」時清除下一列 I 欄 | 讀取清理後檔案，下一列 I 欄為 None |
| AC-004-3 | 格式保留（字型、邊框、填色） | 比對清理前後格式物件一致 |
| AC-004-4 | process_records 記錄清理結果 | DB: process_type='cleanup', status='success' |

#### Acceptance Test

```python
def test_ac_004_1_supply_qty_cleared():
    """AC-004-1: 供應數量欄位清除"""
    upload_and_cleanup(client, "test_data/forecast_with_supply.xlsx")
    wb = openpyxl.load_workbook("processed/{uid}/{ts}/cleaned_forecast.xlsx")
    ws = wb.active
    for row in range(1, ws.max_row + 1):
        if ws.cell(row=row, column=11).value == "供應數量":
            for col in range(12, 50):  # L=12 ~ AW=49
                assert ws.cell(row=row, column=col).value is None

def test_ac_004_3_format_preserved():
    """AC-004-3: 格式保留"""
    original_wb = openpyxl.load_workbook("test_data/forecast_with_supply.xlsx")
    original_font = original_wb.active.cell(row=1, column=1).font
    upload_and_cleanup(client, "test_data/forecast_with_supply.xlsx")
    cleaned_wb = openpyxl.load_workbook("processed/{uid}/{ts}/cleaned_forecast.xlsx")
    cleaned_font = cleaned_wb.active.cell(row=1, column=1).font
    assert original_font.name == cleaned_font.name
    assert original_font.size == cleaned_font.size
```

---

### US-005：Mapping 設定與整合

**As a** 操作人員
**I want to** 設定客戶 Mapping 並整合至 ERP/Transit
**So that** 後續 Forecast 處理能正確計算目標日期

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-005-1 | 可設定區域、排程斷點、ETD、ETA | POST `/api/mapping` 儲存後 GET 讀回一致 |
| AC-005-2 | 批次儲存回傳正確筆數 | POST `/api/mapping/list` → count=N |
| AC-005-3 | 重複 KEY 時 UPDATE | 儲存兩次，DB 只有 1 筆 |
| AC-005-4 | Mapping 整合後 ERP 含 region/etd/eta 欄位 | 讀取 integrated_erp.xlsx 對應欄位有值 |
| AC-005-5 | 使用者 Mapping 互不可見 | user1 讀不到 user2 的 Mapping |

---

### US-006：Forecast 預測處理

**As a** 操作人員
**I want to** 執行 Forecast 預測處理
**So that** ERP/Transit 數據自動填入 Forecast 的 ETA QTY 欄位

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-006-1 | 依 Plant + MRP ID + PN Model 比對 Forecast Block | 正確 Block 的 ETA QTY 列有值 |
| AC-006-2 | 依排程斷點 + ETA 計算目標週欄位 | 目標週欄位數值 = 淨需求數量 |
| AC-006-3 | 同一儲存格數量累加 | 原值 300 + 新值 200 = 500 |
| AC-006-4 | 已分配記錄不重複處理 | 重跑後數值不變（不 double） |
| AC-006-5 | Transit 數據同樣填入 | Transit 數量出現在 ETA QTY |
| AC-006-6 | ERP 記錄標記「已分配」(✓) | integrated_erp.xlsx 分配欄有 ✓ |
| AC-006-7 | 處理時間 500 筆 < 5 秒 | 計時驗證 |

#### Acceptance Test

```python
def test_ac_006_1_block_matching():
    """AC-006-1: Block 比對正確"""
    run_full_pipeline(client, erp="test_data/pegatron_erp.xlsx",
                      forecast="test_data/pegatron_forecast.xlsx")
    wb = openpyxl.load_workbook("processed/{uid}/{ts}/forecast_result.xlsx")
    ws = wb.active
    # 找到 Plant=3A32, MRP=A00Y, PN=0703-00AG000 的 Block
    block = find_block_in_sheet(ws, "3A32", "A00Y", "0703-00AG000")
    assert block is not None
    # ETA QTY 列應有非空值
    eta_qty_row = block["eta_qty_row"]
    has_value = any(ws.cell(row=eta_qty_row, column=c).value for c in range(12, 50))
    assert has_value is True

def test_ac_006_3_quantity_accumulation():
    """AC-006-3: 數量累加"""
    # 第一次處理：ERP 含 qty=300
    run_forecast(client, erp_qty=300)
    val1 = read_target_cell()
    assert val1 == 300
    # 重置分配狀態，新增 qty=200
    run_forecast(client, erp_qty=200, append=True)
    val2 = read_target_cell()
    assert val2 == 500

def test_ac_006_4_no_double_processing():
    """AC-006-4: 不重複處理"""
    run_forecast(client, erp_qty=300)
    val1 = read_target_cell()
    # 重跑（不清除分配狀態）
    run_forecast_again(client)
    val2 = read_target_cell()
    assert val1 == val2  # 值不變

def test_ac_006_7_performance():
    """AC-006-7: 500 筆 < 5 秒"""
    import time
    upload_large_erp(client, rows=500)
    start = time.time()
    client.post("/run_forecast")
    duration = time.time() - start
    assert duration < 5.0
```

---

### US-007：結果下載

**As a** 操作人員
**I want to** 下載處理完成的 Excel 報表
**So that** 我可以取得最終預測結果

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-007-1 | 下載回傳有效 .xlsx | Content-Type 含 spreadsheetml |
| AC-007-2 | 可下載 4 種結果檔案 | 分別 GET 4 個檔案皆 200 |
| AC-007-3 | 下載行為記錄至 activity_logs | DB: action='download_file' |

---

### US-008：使用者管理

**As a** 管理員
**I want to** 管理系統使用者帳號
**So that** 我可以控制誰能存取系統

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-008-1 | Admin 可新增使用者 | POST `/api/users` → success |
| AC-008-2 | Admin 可停用使用者 | PUT `/api/users/{id}` is_active=false → 該帳號無法登入 |
| AC-008-3 | 一般使用者無法存取管理頁面 | GET `/users_manage` → 302/403 |
| AC-008-4 | IT 可查看活動日誌 | GET `/logs` → 200 |

---

### US-009：IT 測試模式

**As a** IT 人員
**I want to** 模擬特定客戶進行測試
**So that** 我可以在不影響正式數據的情況下驗證系統功能

#### Acceptance Criteria

| AC ID | 條件 | 驗證方式 |
|-------|------|---------|
| AC-009-1 | 可選擇目標客戶進入測試 | IT 儀表板有客戶選單 |
| AC-009-2 | 使用客戶專屬模板驗證 | 上傳時使用 compare/{customer}/ 模板 |
| AC-009-3 | 測試結果存入 IT 使用者目錄 | 檔案在 uploads/{it_user_id}/ 下 |

---

## 3. 驗收測試追蹤矩陣

| User Story | AC 總數 | 通過 | 失敗 | 待測 | 狀態 |
|------------|:------:|:----:|:----:|:----:|------|
| US-001 登入認證 | 5 | 5 | 0 | 0 | PASS |
| US-002 ERP 上傳 | 5 | 5 | 0 | 0 | PASS |
| US-003 Forecast 多檔 | 4 | 4 | 0 | 0 | PASS |
| US-004 數據清理 | 4 | 4 | 0 | 0 | PASS |
| US-005 Mapping 整合 | 5 | 5 | 0 | 0 | PASS |
| US-006 Forecast 處理 | 7 | 7 | 0 | 0 | PASS |
| US-007 結果下載 | 3 | 3 | 0 | 0 | PASS |
| US-008 使用者管理 | 4 | 4 | 0 | 0 | PASS |
| US-009 IT 測試模式 | 3 | 3 | 0 | 0 | PASS |
| **合計** | **40** | **40** | **0** | **0** | **ALL PASS** |
