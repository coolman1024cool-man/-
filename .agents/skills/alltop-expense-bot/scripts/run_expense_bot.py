import pandas as pd
import pyautogui
import pyperclip
import time
import sys
import os

print("==================================================")
print("   ALLTOP 經費核銷 RPA 自動化填單機器人 v1.1   ")
print("==================================================")

excel_file = '核銷自動化專用表.xlsx'
if not os.path.exists(excel_file):
    print(f"錯誤：找不到 {excel_file}，請確認檔案放置於同目錄下！")
    input("按 Enter 鍵結束...")
    sys.exit(1)

try:
    df = pd.read_excel(excel_file)
    df = df.fillna('')
except Exception as e:
    print(f"讀取 Excel 失敗: {e}")
    input("按 Enter 鍵結束...")
    sys.exit(1)

total_rows = len(df)
print(f"成功讀取 Excel，共找到 {total_rows} 筆核銷單據。")
print("--------------------------------------------------")

print("\n【操作模式選擇】")
print("1. 步驟一：填寫「主單」核銷明細 (品名與金額)")
print("2. 步驟二：填寫「付款明細」與「發票/憑證明細」")
print("3. 完整填報模式 (依序引導步驟一與步驟二)")
mode = input("請選擇欲執行的步驟 (預設 1): ").strip()
if not mode:
    mode = "1"

print("\n請輸入您要從第幾筆開始填寫？")
print(f"(直接按 Enter 預設從第 1 筆開始，輸入範圍: 1 ~ {total_rows})")
start_idx_str = input(">> ").strip()

start_idx = 1
if start_idx_str.isdigit():
    start_idx = int(start_idx_str)
    if start_idx < 1 or start_idx > total_rows:
        print("輸入超出範圍，將從第 1 筆開始。")
        start_idx = 1

DELAY_AFTER_PASTE = 0.5
TAB_INTERVAL = 0.3

def countdown(sec=3, msg="請切換至瀏覽器點擊對應欄位..."):
    print(f"{msg} (倒數 {sec} 秒)")
    for i in range(sec, 0, -1):
        print(f"{i}...")
        time.sleep(1)

for index, row in df.iterrows():
    curr_num = index + 1
    if curr_num < start_idx:
        continue
    
    item_name = str(row.get('使用說明(品名)', '')).strip()
    amount = str(row.get('申請金額', '')).strip()
    pay_type = str(row.get('付款類別', '付款資料')).strip()
    target_id = str(row.get('受款人代號', '')).strip()
    target_name = str(row.get('受款人姓名', '')).strip()
    tax_rate = str(row.get('所得稅率(%)', 0)).strip()
    inv_type = str(row.get('憑證類別', '免用統一發票')).strip()
    tax_id = str(row.get('廠商統編', '')).strip()
    date_str = str(row.get('單據日期', '')).strip()
    inv_no = str(row.get('發票號碼', '')).strip()

    print("\n" + "="*55)
    print(f"👉 正在處理第 {curr_num} 筆：【{item_name}】 金額: ${amount}")
    print(f"   付款類別: {pay_type} | 受款對象: {target_name} ({target_id}) | 稅率: {tax_rate}%")
    print(f"   憑證類別: {inv_type} | 統編: {tax_id} | 日期: {date_str} | 發票號: {inv_no}")
    print("="*55)

    # 執行步驟一：核銷明細新增
    if mode in ["1", "3"]:
        input(f"\n[步驟 1/2] 請在「核銷明細」點擊【新增】後，按 Enter 準備填入【品名: {item_name}】...")
        countdown(3, "請點擊「使用說明(品名)」輸入框")
        
        pyperclip.copy(item_name)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(DELAY_AFTER_PASTE)
        
        # 移動至申請金額欄位
        pyautogui.press('tab', presses=2, interval=TAB_INTERVAL)
        pyperclip.copy(amount)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(DELAY_AFTER_PASTE)
        print("✅ 已自動填入品名與金額！請確認預算代碼無誤後點擊【確認】。")

    # 執行步驟二：付款明細與發票明細
    if mode in ["2", "3"]:
        input(f"\n[步驟 2/2] 請切換至「付款明細」分頁並點擊【新增】後，按 Enter 準備填入受款資料...")
        countdown(3, "請點擊「人事代號/統一編號」輸入框")
        
        pyperclip.copy(target_id)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(DELAY_AFTER_PASTE)
        print(f"✅ 已填入受款人代號: {target_id} (請點擊「選擇」帶出帳號)")
        
        input("\n帶出帳號後，按 Enter 準備自動填寫「費用說明」與「給付總額」...")
        countdown(3, "請點擊「費用說明」輸入框")
        
        pyperclip.copy(item_name)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(DELAY_AFTER_PASTE)
        
        # 貼上給付總額提示
        pyperclip.copy(amount)
        print(f"✅ 費用說明已填入！剪貼簿已備妥給付總額: ${amount} (若需貼上請至給付總額按 Ctrl+V)")
        
        if pay_type in ["所得資料", "付款兼所得"] or (tax_rate and float(tax_rate) > 0):
            print(f"⚠️ 注意：此筆為所得申報，請確認勾選/填寫所得稅扣繳率為 {tax_rate}%！")

        print("\n--------------------------------------------------")
        print(f"📝 接下來請填報發票/憑證明細：")
        print(f"   憑證類別：{inv_type}")
        print(f"   廠商統編：{tax_id}")
        print(f"   開立日期：{date_str}")
        if inv_no:
            print(f"   發票號碼：{inv_no}")
        print(f"   申報金額：${amount}")
        print("--------------------------------------------------")
        
        input("請點開下方「發票明細 [新增]」彈窗，按 Enter 剪貼簿將為您複製【統編/發票號】...")
        if inv_type == '發票' and inv_no:
            pyperclip.copy(inv_no)
            print(f"✅ 已為您複製發票號碼: {inv_no} (直接 Ctrl+V 貼上)")
        else:
            pyperclip.copy(tax_id)
            print(f"✅ 已為您複製廠商統編: {tax_id} (直接 Ctrl+V 貼上)")

print("\n" + "="*55)
print("🎉 所有選定項目已處理完畢！")
print("請切換至「附件」分頁上傳佐證文件後，點擊 [確認] 送出申請單。")
print("==================================================")
input("按 Enter 鍵關閉視窗...")
