@echo off
chcp 65001 >nul
title ALLTOP 經費核銷自動化機器人
cd /d "%~dp0"
python -X utf8 run_expense_bot.py
pause
