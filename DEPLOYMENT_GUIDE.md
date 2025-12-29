# 1Panel 部署指南 - 解決快取問題

## 問題描述
無痕模式正常但部署到1Panel後出現快取問題，瀏覽器會快取舊版本的靜態資源。

## 解決方案

### 1. 應用層面修改（已完成）

#### 1.1 版本控制機制
- 在 `app.py` 中添加了基於時間戳的版本控制
- 所有靜態資源（CSS、JS）都會附加版本參數 `?v=timestamp`
- HTML頁面設置為不緩存

#### 1.2 緩存策略
- HTML頁面：完全不緩存
- CSS/JS文件：緩存1小時
- 圖片文件：緩存24小時
- 上傳/處理文件：不緩存

### 2. 1Panel 部署配置

#### 2.1 應用配置
1. 在1Panel中創建新的應用
2. 選擇Python環境
3. 設置以下配置：
   ```
   應用名稱: business-forecasting
   端口: 12026
   啟動命令: python app.py
   工作目錄: /path/to/business_forecasting
   ```

#### 2.2 Nginx 配置
1. 在1Panel的網站管理中創建新網站
2. 使用提供的 `nginx.conf` 配置
3. 修改以下路徑為實際路徑：
   - `root /path/to/your/business_forecasting;`
   - `alias /path/to/your/business_forecasting/static/;`
   - `alias /path/to/your/business_forecasting/uploads/;`
   - `alias /path/to/your/business_forecasting/processed/;`

#### 2.3 環境變量設置
在1Panel應用設置中添加：
```
FLASK_ENV=production
FLASK_DEBUG=False
```

### 3. 部署步驟

#### 3.1 上傳代碼
1. 將整個 `business_forecasting` 目錄上傳到服務器
2. 確保目錄權限正確：
   ```bash
   chmod -R 755 /path/to/business_forecasting
   chown -R www-data:www-data /path/to/business_forecasting
   ```

#### 3.2 安裝依賴
在1Panel的終端中執行：
```bash
cd /path/to/business_forecasting
pip install -r requirements.txt
```

#### 3.3 配置Nginx
1. 複製 `nginx.conf` 內容到1Panel的Nginx配置
2. 修改路徑為實際路徑
3. 重載Nginx配置

#### 3.4 啟動應用
在1Panel中啟動應用，確保端口12026正常運行。

### 4. 驗證部署

#### 4.1 檢查版本控制
1. 打開瀏覽器開發者工具
2. 查看Network標籤
3. 確認CSS和JS文件都有 `?v=timestamp` 參數
4. 確認HTML頁面返回 `Cache-Control: no-cache`

#### 4.2 測試快取清除
1. 修改CSS或JS文件
2. 重啟應用（會生成新的時間戳）
3. 刷新頁面，確認使用新版本

### 5. 故障排除

#### 5.1 如果仍有快取問題
1. 檢查瀏覽器是否啟用了強制刷新（Ctrl+F5）
2. 清除瀏覽器緩存
3. 檢查Nginx配置是否正確應用

#### 5.2 檢查日誌
```bash
# 查看應用日誌
tail -f /var/log/1panel/apps/business-forecasting/app.log

# 查看Nginx日誌
tail -f /var/log/nginx/business_forecasting_access.log
tail -f /var/log/nginx/business_forecasting_error.log
```

#### 5.3 手動清除快取
如果問題持續，可以手動清除：
```bash
# 清除Nginx緩存
sudo nginx -s reload

# 重啟應用
sudo systemctl restart business-forecasting
```

### 6. 性能優化建議

1. **CDN配置**：考慮使用CDN來分發靜態資源
2. **Gzip壓縮**：在Nginx中啟用Gzip壓縮
3. **HTTP/2**：啟用HTTP/2支持
4. **SSL證書**：配置HTTPS證書

### 7. 監控和維護

1. 定期檢查應用運行狀態
2. 監控磁盤空間（上傳文件會佔用空間）
3. 定期清理舊的處理結果文件
4. 備份重要配置文件

## Chrome 瀏覽器阻擋功能

### 功能說明
系統已添加Chrome瀏覽器檢測和強制阻擋功能：

1. **自動檢測**：頁面載入時自動檢測用戶使用的瀏覽器
2. **Chrome強制阻擋**：如果檢測到Google Chrome瀏覽器，會顯示無法關閉的阻擋彈窗
3. **Edge直接開啟**：提供「用 Edge 開啟」按鈕，直接啟動Edge瀏覽器開啟應用
4. **Edge下載**：提供Microsoft Edge瀏覽器下載連結作為備選方案
5. **強制性設計**：彈窗無法關閉，確保用戶必須使用Edge瀏覽器

### 測試方法
1. 使用Chrome瀏覽器打開應用，應該會看到無法關閉的阻擋彈窗
2. 點擊「用 Edge 開啟」按鈕，應該會自動啟動Edge瀏覽器
3. 使用Edge瀏覽器打開應用，應該正常顯示
4. 可以打開 `test_browser.html` 測試瀏覽器檢測功能

### 自定義選項
如果需要修改阻擋行為，可以編輯模板文件中的JavaScript代碼：
- 修改檢測邏輯：`detectBrowser()` 函數
- 修改彈窗內容：模態框HTML結構
- 修改樣式：CSS中的 `.modal` 相關樣式

## LibreOffice 安裝說明（重要）

系統使用 LibreOffice 處理 Excel 檔案（.xls/.xlsx），需要在 Linux 伺服器上安裝 LibreOffice。

### Ubuntu/Debian 安裝

```bash
# 更新套件列表
sudo apt-get update

# 安裝 LibreOffice Calc 和無頭模式支援
sudo apt-get install -y libreoffice-calc libreoffice-headless

# 驗證安裝
libreoffice --version
```

### CentOS/RHEL 安裝

```bash
# 安裝 LibreOffice
sudo yum install -y libreoffice-calc

# 或使用 dnf（CentOS 8+）
sudo dnf install -y libreoffice-calc

# 驗證安裝
libreoffice --version
```

### 1Panel 環境安裝

在 1Panel 終端中執行：

```bash
# Ubuntu/Debian 系統
apt-get update && apt-get install -y libreoffice-calc libreoffice-headless

# CentOS 系統
yum install -y libreoffice-calc
```

### 驗證 LibreOffice 安裝

安裝完成後，執行以下命令驗證：

```bash
# 檢查版本
libreoffice --version

# 測試轉換功能（可選）
echo "test" > /tmp/test.txt
libreoffice --headless --convert-to xlsx /tmp/test.txt --outdir /tmp/
ls /tmp/test.xlsx
```

### 常見問題

1. **權限問題**：確保運行應用的用戶有執行 libreoffice 的權限
2. **路徑問題**：如果 libreoffice 不在 PATH 中，可能需要使用完整路徑
3. **字體問題**：如果 Excel 中有中文字體顯示異常，安裝中文字體：
   ```bash
   # Ubuntu/Debian
   sudo apt-get install fonts-wqy-zenhei fonts-wqy-microhei

   # CentOS
   sudo yum install wqy-zenhei-fonts wqy-microhei-fonts
   ```

## 注意事項

1. 每次代碼更新後，版本號會自動更新，無需手動操作
2. 如果修改了靜態資源，建議重啟應用以確保版本號更新
3. 生產環境建議關閉Flask的debug模式
4. 定期檢查和清理uploads和processed目錄中的舊文件
5. Chrome阻擋功能會影響用戶體驗，請根據實際需求決定是否啟用
6. **必須安裝 LibreOffice**：系統需要 LibreOffice 來處理 Excel 檔案，特別是 .xls 格式的公式保留
