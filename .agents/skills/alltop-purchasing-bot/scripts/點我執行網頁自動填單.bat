@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================
echo     ALLTOP 網頁自動填單機器人
echo ==========================================
python run_bot.py
