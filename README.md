# FORECAST 數據處理系統

一個智能的預測數據整合與分析平台，提供完整的數據處理流程。

## 🌟 功能特色

- **多客戶支援**: 廣達 (Quanta)、台達 (Delta) 獨立處理流程
- **文件上傳**: 支援ERP淨需求文件和Forecast文件上傳
- **數據清理**: 自動清理Forecast文件中的特定數據
- **映射整合**: 可視化配置客戶簡稱與地區的映射關係
- **FORECAST處理**: 超高速批量處理，性能提升20倍
- **累加邏輯**: 支援數據累加，避免覆蓋問題
- **結果下載**: 一鍵下載所有處理結果

### 🆕 台達 (Delta) 專屬流程
- **8 檔合併**: 自動合併 8 個不同 PLANT 格式 Forecast (Ketwadee/Kanyanat/Weeraya/India IAI1/PSW1+CEW1/MWC1+IPC1/NBQ1/SVC1+PWC1)
- **匯總格式模式 (方案二)**: 固定 26 日期欄 = `PASSDUE + 16 週 + 9 月`
- **動態 W1 起點**: 以來源檔最早的週一為 W1，避免 PASSDUE 被誤折疊
- **多對一累加**: 同月份的多個週末日期自動折疊累加至月份欄
- **Transit + ERP 回填**: 將 Transit 和 ERP 淨需求回填到合併後的 Forecast

## 🎨 設計風格

採用沉穩文清風格設計：
- 漸層色彩搭配
- 圓角卡片設計
- 流暢動畫效果
- 響應式布局
- 直觀的進度指示器

## 🚀 快速開始

### 1. 安裝依賴

```bash
pip install -r requirements.txt
```

### 2. 啟動應用

```bash
python app.py
```

### 3. 訪問系統

打開瀏覽器訪問：`http://localhost:12058`

## 📋 使用流程

### 第一階段：文件上傳
1. 上傳 `20250924 廣達淨需求.xlsx` (ERP文件)
2. 上傳 `ForecastDataFile_ALL-0923.xlsx` (Forecast文件)

### 第二階段：數據清理
- 自動清理Forecast文件中的特定數據
- 當K欄位是"供應數量"時，清空L~AW欄位
- 當I欄位有"庫存數量"時，清空下一列I欄位

### 第三階段：映射整合
1. 點擊"配置映射表"按鈕
2. 為每個客戶簡稱配置：
   - 客戶需求地區
   - 排程出貨日期斷點
   - ETD
   - ETA
3. 保存配置並返回主頁
4. 點擊"開始整合"執行映射

### 第四階段：FORECAST處理
- 執行智能預測數據填寫
- 支援累加邏輯
- 超高速批量處理

### 第五階段：結果下載
- 下載清理後的Forecast文件
- 下載整合後的ERP文件
- 下載FORECAST處理結果

---

## 🏭 台達 (Delta) 專屬使用流程

### Step 1 — 上傳 8 個 Forecast 檔案
| 類別 | 檔名範例 | PLANT |
|------|---------|-------|
| Ketwadee | `PSBG PSB5- Ketwadee...xlsx` | PSB5 |
| Kanyanat | `PSBG PSB7_Kanyanat.S...xlsx` | PSB7 |
| Weeraya | `PSBG PSB7-Weeraya...xlsx` | PSB7 |
| India IAI1 | `PSBG (India IAI1&UPI2&DFI1 DIODES)...xlsx` | IAI1/UPI2 |
| PSW1+CEW1 | `PSBG PSW1+CEW1-...xlsx` | PSW1/CEW1 |
| MWC1+IPC1 | `強茂 MWC1IPC1 MRP...xlsx` | MWC1/IPC1 |
| NBQ1 | `NBQ1.xlsx` | NBQ1 |
| SVC1+PWC1 Diode&MOS | `MRP(SVC1PWC1 DIODE&MOS)...xlsx` | SVC1/PWC1 |

系統自動合併為匯總格式 (方案二)：
- K~Z 欄: **W1~W16** 週 (以來源檔最早週一起算)
- AA~AI 欄: **9 個月份** (W16 次週起算)
- PASSDUE 欄: 來自源檔明確標籤，不會從日期折入

### Step 2 — 上傳 ERP + Transit
- **ERP 淨需求**: `0408-上午淨需求 (台達).xlsx` (含 "PJOMR006 for SBU" 工作表)
- **Transit 在途**: 可選，若無則跳過

### Step 3 — 映射整合
系統依據 DB 的 Delta 客戶映射表 (11 筆) 填入：
- 客戶需求地區 / 斷點 / ETD / ETA
- Forecast C/D 欄 (客戶簡稱 + 送貨地點)

### Step 4 — Transit + ERP 回填
- **Transit**: 依 (客戶簡稱, 送貨地點, 料號) 回填在途數量
- **ERP**: 依 (客戶簡稱, 送貨地點, 料號) 回填淨需求至對應週/月欄位
- 依 ETA 自動歸入 W1~W16 或月份欄位

## 🛠️ 技術架構

### 後端
- **Flask**: Web框架
- **pandas**: 數據處理
- **openpyxl**: Excel文件操作

### 前端
- **HTML5**: 語義化標記
- **CSS3**: 現代化樣式
- **JavaScript**: 交互邏輯
- **Font Awesome**: 圖標庫

### 文件結構
```
business_forecasting/
├── app.py                          # Flask主應用
├── ultra_fast_forecast_processor.py # 廣達 FORECAST處理核心
├── delta_forecast_processor.py     # 台達 8檔合併 + 方案二匯總格式
├── delta_forecast_step4.py         # 台達 Transit + ERP 回填
├── database.py                     # MySQL 使用者 / 映射表
├── requirements.txt                # 依賴包列表
├── README.md                      # 說明文檔
├── docs/
│   └── delta_meeting_notes.md     # 台達會議紀錄
├── templates/                     # HTML模板
│   ├── index.html                 # 主頁面
│   └── mapping.html               # 映射配置頁面
├── static/                        # 靜態資源
│   ├── css/
│   │   └── style.css              # 樣式文件
│   └── js/
│       ├── main.js                # 主頁面邏輯
│       └── mapping.js             # 映射頁面邏輯
├── uploads/                       # 上傳文件目錄
└── processed/                     # 處理結果目錄
```

## 📊 處理邏輯

### 數據清理邏輯
- 掃描每行數據
- 識別特定欄位值
- 清空指定範圍的數據
- 保持原始格式和架構

### 映射整合邏輯
- 根據客戶簡稱匹配
- 添加映射欄位
- 按排程出貨日期排序
- 保持數據完整性

### FORECAST處理邏輯
- 識別數據塊
- 匹配客戶料號和地區
- 計算ETA目標日期
- 轉換數值單位
- 累加相同位置的值
- 批量寫入結果

## 🔧 配置說明

### 環境要求
- Python 3.8+
- 現代瀏覽器
- 足夠的內存處理大型Excel文件

### 性能優化
- 使用批量openpyxl操作
- 內存中處理減少I/O
- 索引優化查找速度
- 累加邏輯避免重複處理

## 📝 注意事項

1. **文件格式**: 僅支援 `.xlsx` 和 `.xls` 格式
2. **文件大小**: 建議單個文件不超過50MB
3. **欄位名稱**: 系統會自動識別包含關鍵字的欄位
4. **數據備份**: 建議在處理前備份原始文件
5. **瀏覽器支援**: 建議使用Chrome、Firefox或Edge

## 🐛 故障排除

### 常見問題

**Q: 上傳文件失敗**
A: 檢查文件格式是否為Excel，文件是否損壞

**Q: 找不到客戶簡稱欄位**
A: 確保欄位名稱包含"客戶"和"簡稱"關鍵字

**Q: 映射配置保存失敗**
A: 檢查所有必填欄位是否已填寫

**Q: FORECAST處理失敗**
A: 確保已完成前兩個階段的處理

### 日誌查看
系統會在控制台輸出詳細的處理日誌，包括：
- 文件載入狀態
- 數據處理進度
- 錯誤信息
- 性能統計

## 📞 技術支援

如有問題或建議，請聯繫技術支援團隊。

---

© 2024 FORECAST 數據處理系統. 智能預測，精準分析.
