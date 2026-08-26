
import pandas as pd
import pyautogui
import pyperclip
import time
import sys

print("正在讀取 Excel 資料庫...")
try:
    df = pd.read_excel('採購自動化專用表.xlsx')
    df = df.fillna('')
except Exception as e:
    print(f"讀取 Excel 失敗: {e}")
    sys.exit(1)

print(f"讀取成功！共找到 {len(df)} 筆資料。")
print("==================================================")
print("【超高效率 RPA 機器人 - 慢速安全版】")

# 放慢動作設定
DELAY_AFTER_PASTE = 0.6  # 貼上後的停頓時間 (秒)
TAB_INTERVAL = 0.4       # 連按 Tab 之間的間隔時間 (秒)

for index, row in df.iterrows():
    input(f"\n▶ 準備填寫 [{row['品名']}]... \n(請按 Enter 開始，然後馬上用滑鼠點擊網頁的「品名」欄位)")
    print("倒數 4 秒，請切換回網頁點擊品名欄...")
    for i in range(4, 0, -1):
        print(f"{i}...")
        time.sleep(1)
    
    # 1. 品名
    pyperclip.copy(str(row['品名']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    # 2. 跳過優先序，來到規格 (Tab x2)
    pyautogui.press('tab', presses=2, interval=TAB_INTERVAL)
    pyperclip.copy(str(row['規格']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    # 3. 用途及說明 (Tab x1)
    pyautogui.press('tab', interval=TAB_INTERVAL)
    pyperclip.copy(str(row['用途及說明']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    # 4. 計量單位 (Tab x1)
    pyautogui.press('tab', interval=TAB_INTERVAL)
    pyperclip.copy(str(row['計量單位']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    # 5. 數量 (Tab x1)
    pyautogui.press('tab', interval=TAB_INTERVAL)
    pyperclip.copy(str(row['數量']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    # 6. 單價 (Tab x1)
    pyautogui.press('tab', interval=TAB_INTERVAL)
    pyperclip.copy(str(row['單價']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    # 7. 到達確認按鈕 (Tab x2)
    pyautogui.press('tab', presses=2, interval=TAB_INTERVAL)
    time.sleep(0.5)
    
    # 自動按下確認 (Enter)
    pyautogui.press('enter')
    
    print(f"完成！第 {index+1} 筆已自動填寫並送出！")

print("\n🎉 所有的資料都已經填寫完畢囉！")
