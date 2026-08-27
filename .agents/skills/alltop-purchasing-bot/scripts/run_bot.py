import pandas as pd
import pyautogui
import pyperclip
import time
import sys

print("正在讀取 Excel 資料中...")
try:
    df = pd.read_excel('採購自動化專用表.xlsx')
    df = df.fillna('')
except Exception as e:
    print(f"讀取 Excel 失敗: {e}")
    sys.exit(1)

total_rows = len(df)
print(f"讀取成功，找到 {total_rows} 筆採購項目。")
print("==================================================")
print("啟動 ALLTOP RPA 機器人")

print("\n請輸入您要從第幾筆開始填寫？")
print(f"(直接按 Enter 預設從第 1 筆開始，輸入範圍: 1 ~ {total_rows})")
start_idx_str = input(">> ").strip()

start_idx = 1
if start_idx_str.isdigit():
    start_idx = int(start_idx_str)
    if start_idx < 1 or start_idx > total_rows:
        print("輸入超出範圍，將從第 1 筆開始。")
        start_idx = 1

DELAY_AFTER_PASTE = 0.6  
TAB_INTERVAL = 0.4       

for index, row in df.iterrows():
    current_item_num = index + 1
    if current_item_num < start_idx:
        continue
        
    input(f"\n準備填寫第 {current_item_num} 筆: [{row['品名']}]... \n(請按 Enter 繼續，然後馬上用滑鼠點擊網頁上的「品名」欄位)")
    print("倒數 4 秒，請切換至網頁點擊欄位...")
    for i in range(4, 0, -1):
        print(f"{i}...")
        time.sleep(1)
    
    pyperclip.copy(str(row['品名']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    pyautogui.press('tab', presses=2, interval=TAB_INTERVAL)
    pyperclip.copy(str(row['規格']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    pyautogui.press('tab', interval=TAB_INTERVAL)
    pyperclip.copy(str(row['用途及說明']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    pyautogui.press('tab', interval=TAB_INTERVAL)
    pyperclip.copy(str(row['計量單位']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    pyautogui.press('tab', interval=TAB_INTERVAL)
    pyperclip.copy(str(row['數量']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    pyautogui.press('tab', interval=TAB_INTERVAL)
    pyperclip.copy(str(row['單價']))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(DELAY_AFTER_PASTE)
    
    pyautogui.press('tab', presses=2, interval=TAB_INTERVAL)
    time.sleep(0.5)
    
    pyautogui.press('enter')
    
    print(f"完成！第 {current_item_num} 筆已自動填寫並送出。")

print("\n所有資料均已填寫完畢！")
