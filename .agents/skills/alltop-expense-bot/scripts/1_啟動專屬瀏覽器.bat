@echo off
title ALLTOP Chrome 自動化啟動器
echo ========================================================
echo   正在以「自動化接管模式」啟動 Chrome 瀏覽器...
echo ========================================================
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebugProfile" "https://pags.mitust.edu.tw/ALLTOP/system.php"
echo.
echo [OK] 瀏覽器已開啟！
echo 請在開啟的 Chrome 中正常登入 ALLTOP 系統，並進入「02040 核銷申請單」頁面。
echo 準備好後，即可隨時雙擊執行【2_執行核銷全自動機器人.bat】！
echo.
pause
