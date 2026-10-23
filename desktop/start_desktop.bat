@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在啟動 MOPS 金融股月自結損益彙整系統 (Windows 桌面版)...
start "" pythonw desktop_app.py
if errorlevel 1 (
    python desktop_app.py
)
