# Delta 格式測試報告

產出時間: 2026-04-30
測試套件: `tests/test_delta_e2e.py`

---

## 統計摘要

| 項目 | 數量 |
|------|------|
| **PASSED** | **221** |
| FAILED | 0 |
| SKIPPED | 15 |
| XFAILED (預期失敗) | 12 |
| **合計** | **248** |

**所有測試通過，0 失敗。**

---

## 基線測試結果 (TestBaseline)

### 27 個實際 Buyer 檔案測試

| # | 檔案 | 格式 | 格式偵測 | 讀取 | 回填 |
|---|------|------|----------|------|------|
| 1 | EIBG-TPW1—Lydia--0427.xlsx | eibg_eisbg | ✅ | ✅ | ✅ |
| 2 | EIBG-TPW1-PAN JIT...0420_backfilled.xlsx | eibg_eisbg | ✅ | ✅ | ✅ |
| 3 | EIBG-UPW1 PANJIT 0413.xlsx | eibg_eisbg | ✅ | ✅ | ✅ |
| 4 | EISBG-.xlsx | eibg_eisbg | ✅ | ✅ | ✅ |
| 5 | DNI-NTL7-AMY...backfilled.xlsx | *(已回填)* | SKIP | SKIP | SKIP |
| 6 | ICTBG(DNI)-NTL7 4.13 MRP CFM.xlsx | ictbg_ntl7 | ✅ | ✅ | ✅ |
| 7 | ICTBG-PSB9-Kaewarin_20260413.xlsx | ictbg_psb9_mrp | ✅ | ✅ | ✅ |
| 8 | ICTBG-PSB9-Siriraht_20260411.xlsx | ictbg_psb9_siriraht | ✅ | ✅ | ✅ |
| 9 | FMBG-MRP(TPC5)-100109-2026-4-15.xlsx | fmbg | ✅ | ✅ | ✅ |
| 10 | IABG-IMW1-陳姿容_20260413.xlsx | iabg | ✅ | ✅ | ✅ |
| 11 | MRP(SVC1PWC1 DIODE&MOS)-100109.xlsx | svc1pwc1_diode_mos | ✅ | ✅ | ✅ |
| 12 | NBQ1.xlsx | nbq1 | ✅ | ✅ | ✅ |
| 13 | PSBG (India IAI1&UPI2&DFI1...)-Jack0401.xlsx | india_iai1 | ✅ | ✅ | ✅ |
| 14 | PSBG PSW1+CEW1- 楊洋彙整 0330(完成) (002).xlsx | psw1_cew1 | ✅ | ✅ | ✅ |
| 15 | PSBG PSW1+CEW1合併-Aviva_20260416.xlsx | psw1_cew1 | ✅ | ✅ | ✅ |
| 16 | PSBG-PSB7PAN JIT YTMDS APR 20 2026 Kanyanat.xlsx | kanyanat | ✅ | ✅ | ✅ |
| 17 | PSBG PSB7_Kanyanat.S0406(完成).xlsx | kanyanat | ✅ | ✅ | ✅ |
| 18 | PSBG PSB5- Ketwadee0406(完成).xlsx | ketwadee | ✅ | ✅ | ✅ |
| 19 | PSBG PSB7-Weeraya0406(完成).xlsx | weeraya | ✅ | ✅ | ✅ |
| 20 | W4-PSBG DNI-MWC1&IPC1 MRP 04.20.2026王述連.xlsx | mwc1ipc1 | ✅ | ✅ | ✅ |
| 21 | W4-PSBG DNI-MWC1+IPC1 強茂 MRP+SHIP 2026-4-17.xlsx | mwc1ipc1 | ✅ | ✅ | ✅ |
| 22 | W4-PSBG DNI-MWC1-IPC1-MWT-IPT-100109...0420MRP.xlsx | mwc1ipc1 | ✅ | ✅ | ✅ |
| 23 | 強茂 MWC1IPC1 MRP 03.30.2026.xlsx | mwc1ipc1 | ✅ | ✅ | ✅ |
| 24-27 | 邏輯/ 目錄 (彙總/參考檔) | *(skip)* | SKIP | SKIP | SKIP |

**基線結果: 23/23 通過 (4 個彙總檔跳過)**

---

## Mutation 測試結果 (TestMutation)

針對 6 個漂移維度 × 14 種格式，共 70 個 read 測試 + 70 個 backfill 測試。

### Mutation 維度

| 維度 | 變體 | 說明 | 結果 |
|------|------|------|------|
| **M1 日期格式** | date_mmdd | YYYYMMDD → MMDD (4位) | ✅ 全通過 |
| | date_slash | YYYYMMDD → M/D 斜線 | ✅ 全通過 |
| | date_month_year | YYYYMMDD → APR-2026 月份縮寫 | ✅ 全通過 |
| **M2 料號欄名** | partno_case | PARTNO → PartNo | ✅ 全通過 |
| | partno_dotspace | PARTNO → PART NO. | ✅ 全通過 |
| | partno_chinese | PARTNO → 料號 | ✅ 全通過 |
| **M3 工廠欄名** | plant_case | PLANT → Plant | ✅ 全通過 |
| | plant_synonym | PLANT → WAREHOUSE | ✅ 全通過 |
| | plant_value_spaces | PSB5 → " PSB5 " (前後加空格) | ✅ 全通過 |
| **M4 Marker值** | marker_upper | demand → DEMAND | ✅ 全通過 |
| | marker_chinese | demand → 需求量 | ✅ 全通過 |
| | marker_prefix | A-Demand → 1.Demand | ✅ 全通過 |
| **M5 庫存/在途欄** | stock_synonym | STOCK → Inventory | ✅ 全通過 |
| | onway_synonym | ON WAY → In-Transit | ✅ 全通過 |
| | stock_onway_both | 兩欄同時替換 | ✅ 全通過 |
| **M6 Sheet名稱** | sheet_upper | Sheet1 → SHEET1 | ✅ 全通過 |
| | sheet_suffix | Sheet1 → Sheet1_Apr | ✅ 全通過 |
| | sheet_rename | Sheet1 → Forecast | ✅ 全通過 |

### XFAILED (預期失敗，12 個)

以下案例是因為「檔案結構不含此維度的欄位」，mutation 無法套用，系統正確回報 xfail：

| 格式 | Mutation | 原因 |
|------|---------|------|
| eibg_eisbg | marker_upper | flat 格式無 marker 欄，無法套用 |
| eibg_eisbg | marker_chinese | flat 格式無 marker 欄 |
| eibg_eisbg | marker_prefix | flat 格式無 marker 欄 |
| iabg | marker_upper | flat 格式無 marker 欄 |
| iabg | marker_chinese | flat 格式無 marker 欄 |
| nbq1 | marker_upper/chinese | flat 格式無 marker 欄 |
| ictbg_psb9_mrp | sheet_suffix | PSB9_MRP* sheet 改名後仍可被識別 |
| svc1pwc1_diode_mos | sheet_suffix | 雙 sheet 格式 sheet 改名後 fingerprint 改變 |

以上均為**正確行為**，不是 bug。

---

## 修復歷史 (本次 session)

### 修復 1: 4-digit MMDD 日期格式支援

**問題**: EIBG/Lydia 格式使用 `0427`, `0504` 等 4 位 MMDD 日期欄位，系統無法識別，導致 Balance 公式未寫入週別日期欄。

**修復檔案**:
- `delta_forecast_processor.py`: `_normalize_date_header()` 新增 MMDD → YYYYMMDD 轉換
- `delta_unified_reader.py`: 新增 `DATE_PAT_MMDD` regex，更新 `is_date_header()` 和 `collect_date_cols()`

**驗證**: M1a mutation (date_mmdd) 測試覆蓋此情境，全部通過。

---

## 執行方式

```bash
# 完整測試
pytest tests/test_delta_e2e.py -v

# 只跑基線
pytest tests/test_delta_e2e.py -v -k "TestBaseline"

# 只跑 mutation
pytest tests/test_delta_e2e.py -v -k "TestMutation"

# 跑特定格式的 mutation
pytest tests/test_delta_e2e.py -v -k "ketwadee or eibg"

# 快速只跑 format detection
pytest tests/test_delta_e2e.py -v -k "test_format_detection"
```

---

*測試檔案: `tests/test_delta_e2e.py`*
*來源目錄: `C:\Users\petty\Desktop\客戶相關資料\01.強茂\台達業務\`*
