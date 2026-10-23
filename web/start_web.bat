@echo off
chcp 65001 >nul
echo Starting MOPS Financial Stocks Monthly Self-Profit Web Server...
start "" "http://127.0.0.1:8765"
python "%~dp0web_app.py" 8765
pause
