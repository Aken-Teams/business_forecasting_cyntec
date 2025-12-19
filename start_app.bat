@echo off
echo 啟動 FORECAST 數據處理系統...
echo.
echo 正在檢查Python環境...
python --version
echo.
echo 正在安裝依賴包...
pip install -r requirements.txt
echo.
echo 正在啟動Flask應用...
echo 請在瀏覽器中訪問: http://localhost:12026
echo 按 Ctrl+C 停止服務
echo.
python app.py
pause
