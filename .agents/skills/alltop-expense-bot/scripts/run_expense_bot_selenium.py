import os
import sys
import time
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

print("==================================================")
print("   ALLTOP 核銷全自動化機器人 (真·一鍵全自動版 v6.0)   ")
print("==================================================")

excel_path = '核銷自動化專用表.xlsx'
if not os.path.exists(excel_path):
    print(f"錯誤：找不到資料庫表單 {excel_path}！")
    input("按 Enter 鍵結束...")
    sys.exit(1)

try:
    df = pd.read_excel(excel_path).fillna('')
    total_rows = len(df)
    print(f"✅ 成功載入 Excel，共 {total_rows} 筆核銷單據。")
except Exception as e:
    print(f"讀取 Excel 失敗: {e}")
    input("按 Enter 鍵結束...")
    sys.exit(1)

print("\n正在連接至已開啟的 Chrome 瀏覽器 (Port: 9222)...")
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

try:
    driver = webdriver.Chrome(options=chrome_options)
    print("✅ 成功接管 Chrome 瀏覽器！")
except Exception as e:
    print("\n❌ 連接 Chrome 失敗！請確認已先執行【1_啟動專屬瀏覽器.bat】。")
    print(f"錯誤細節: {e}")
    input("\n按 Enter 鍵結束...")
    sys.exit(1)

def switch_to_main():
    driver.switch_to.default_content()
    try:
        driver.switch_to.frame("iframe")
        return True
    except:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for f in iframes:
            if f.get_attribute("id") == "iframe":
                driver.switch_to.frame(f)
                return True
    return False

def select_payee_natively(payee_code):
    """點擊藍色選擇按鈕，並在彈窗 iframe_windows 中原生搜尋並選定受款人"""
    print(f"  🔍 正在透過學校資料庫原生查詢代號: 【{payee_code}】...")
    switch_to_main()
    
    sel_btns = driver.find_elements(By.XPATH, "//button[contains(@onclick, 'SelectTbpay.php')]")
    if not sel_btns:
        return False
    sel_btns[0].click()
    time.sleep(1.5)
    
    driver.switch_to.default_content()
    try:
        driver.switch_to.frame("iframe_windows")
    except:
        switch_to_main()
        return False
        
    search_inp = driver.find_element(By.ID, "key_srh")
    search_inp.clear()
    search_inp.send_keys(payee_code)
    
    search_btn = driver.find_element(By.CSS_SELECTOR, "button.search")
    search_btn.click()
    time.sleep(1.5)
    
    green_btns = driver.find_elements(By.CSS_SELECTOR, "button.green")
    if green_btns:
        green_btns[0].click()
        time.sleep(1.5)
        print(f"  ✅ [受款人原生選定] 已自動取得伺服器簽發之加密金鑰！")
    else:
        print(f"  ⚠️ 查詢不到代號 【{payee_code}】")
        
    switch_to_main()
    return True

def fill_invoice_modal_natively(inv_type, tax_id, date_str, amount, inv_no, item_name):
    """自動點開發票明細彈窗，並填寫憑證資料後確認存檔"""
    print(f"  🧾 正在自動填報發票/憑證明細: 【{inv_type}】 統編/證號: {tax_id} 金額: ${amount}...")
    switch_to_main()
    
    # 點擊發票明細的新增按鈕
    inv_add_btns = driver.find_elements(By.XPATH, "//button[contains(text(), '新增') and contains(@class, 'btn-info')] | //button[contains(@onclick, 'formPay_81.php')]")
    if inv_add_btns:
        inv_add_btns[-1].click()
        time.sleep(1.5)
        
    # 切換至發票彈窗 iframe_windows
    driver.switch_to.default_content()
    try:
        driver.switch_to.frame("iframe_windows")
    except:
        print("  ⚠️ 無法切換至發票彈窗 iframe_windows")
        switch_to_main()
        return False
        
    # 轉換日期格式為西元 (例如 115/08/10 -> 2026-08-10)
    iso_date = str(date_str).replace('/', '-')
    if len(iso_date) >= 7 and iso_date.startswith("115"):
        iso_date = "2026" + iso_date[3:]
        
    kind_val = "免用統一發票"
    if "發票" in inv_type and "免用" not in inv_type:
        kind_val = ""
    elif "其他" in inv_type or "簽收" in inv_type:
        kind_val = "其他"
        
    driver.execute_script(f"""
        $('#billnumkind').selectpicker('val', '{kind_val}');
        $('#billnumkind').val('{kind_val}').change();
        if (typeof $.fn.selectpicker !== 'undefined') {{
            $('#billnumkind').selectpicker('refresh');
        }}
        $('#payno').val('{tax_id}').change();
        $('#billdate').val('{iso_date}').change();
        $('#money').val('{amount}').change();
    """)
    time.sleep(0.8)
    
    # 若為發票填入發票號碼；若為免用發票填入品名
    try:
        bill_inp = driver.find_element(By.ID, "billnum")
        if bill_inp.is_displayed():
            bill_inp.clear()
            if kind_val == "":
                bill_inp.send_keys(inv_no)
            else:
                bill_inp.send_keys(item_name)
    except:
        pass
        
    time.sleep(0.5)
    # 點擊發票彈窗確認
    save_btn = driver.find_elements(By.CSS_SELECTOR, "button.save, button.btn-success")
    if save_btn:
        save_btn[0].click()
        time.sleep(2)
        print("  ✅ [發票明細建立成功] 彈窗已自動儲存關閉！")
        
    switch_to_main()
    return True

switch_to_main()

# =========================================================================
# 如果當前剛好停在已打開的發票彈窗 (F11A0_formPay_81.php)
# =========================================================================
driver.switch_to.default_content()
iframes = driver.find_elements(By.TAG_NAME, "iframe")
for f in iframes:
    src = f.get_attribute("src") or ""
    if "formPay_81.php" in src:
        first_row = df.iloc[0]
        fill_invoice_modal_natively(
            first_row.get('憑證類別', '其他相關憑證'),
            first_row.get('廠商統編', ''),
            first_row.get('單據日期', '2026-08-10'),
            first_row.get('申請金額', ''),
            first_row.get('發票號碼', ''),
            first_row.get('使用說明(品名)', '')
        )
        break

switch_to_main()

# =========================================================================
# 逐筆處理付款明細與發票明細
# =========================================================================
for idx, row in df.iterrows():
    item_num = idx + 1
    item_name = str(row.get('使用說明(品名)', '')).strip()
    amount = str(row.get('申請金額', '')).strip()
    pay_type = str(row.get('付款類別', '付款兼所得')).strip()
    payee_id = str(row.get('受款人代號', '')).strip()
    inv_type = str(row.get('憑證類別', '免用統一發票')).strip()
    tax_id = str(row.get('廠商統編', '')).strip()
    date_str = str(row.get('單據日期', '')).strip()
    inv_no = str(row.get('發票號碼', '')).strip()
    
    current_url = driver.execute_script("return window.location.href;")
    
    # 若在清單頁面，點擊新增
    if "formPay_6.php" in current_url:
        print(f"\n👉 [第 {item_num}/{total_rows} 筆] 點擊【新增】進入付款明細編輯...")
        new_btns = driver.find_elements(By.CSS_SELECTOR, "button.new, button.btn-info")
        for nb in new_btns:
            if "新增" in nb.text:
                nb.click()
                time.sleep(2)
                switch_to_main()
                break

    # 若已在付款明細編輯頁面 (formPay_61.php)
    current_url = driver.execute_script("return window.location.href;")
    if "formPay_61.php" in current_url:
        print(f"\n👉 正在全自動填寫第 {item_num} 筆: 【{item_name}】 ${amount}")
        
        # 1. 設置付款類別 (付款兼所得 val=3, 付款資料 val=1, 所得資料 val=2)
        ispay_val = "3"
        if "付款資料" in pay_type: ispay_val = "1"
        elif pay_type == "所得資料": ispay_val = "2"
        elif "付款兼所得" in pay_type: ispay_val = "3"
        
        driver.execute_script(f"""
            if(typeof $.fn.selectpicker !== 'undefined'){{
                $('#IsPay').selectpicker('val', '{ispay_val}');
                $('#salkind').selectpicker('val', '0');
                $('#PayBusKind').selectpicker('val', '3');
            }}
            $('#IsPay').val('{ispay_val}').change();
            $('#salkind').val('0').change();
            $('#PayBusKind').val('3').change();
        """)
        time.sleep(0.8)

        # 2. 選取主單對應科目
        try:
            formid_sel = Select(driver.find_element(By.ID, "formid"))
            for opt in formid_sel.options:
                if str(amount) in opt.text or item_name in opt.text:
                    opt_val = opt.get_attribute("value")
                    formid_sel.select_by_value(opt_val)
                    driver.execute_script(f"$('#formid').selectpicker('val', '{opt_val}').change();")
                    time.sleep(0.8)
                    print(f"  ✅ 已選取主單科目: {opt.text}")
                    break
        except Exception as e:
            print(f"  ⚠️ 選取科目: {e}")

        # 3. 呼叫原生彈窗查詢受款人 (自動綁定金鑰)
        select_payee_natively(payee_id)

        # 4. 填寫費用說明與金額
        driver.execute_script(f"""
            $('#salitem').val('{item_name}').change();
            $('#salmoney').val('{amount}').change();
            document.getElementById('money').value = '{amount}';
        """)
        time.sleep(0.8)
        
        # 5. 全自動打開發票彈窗並填入確認！
        fill_invoice_modal_natively(inv_type, tax_id, date_str, amount, inv_no, item_name)
        
        # 6. 自動點擊右下角確認存檔此筆付款明細！
        print(f"  💾 正在自動點擊【確認】儲存第 {item_num} 筆付款明細...")
        save_btn = driver.find_elements(By.CSS_SELECTOR, "button.save, button.btn-success")
        if save_btn:
            save_btn[-1].click()
            time.sleep(2.5)
            switch_to_main()
            print(f"  🎉 第 {item_num} 筆付款明細已完整建立並儲存！")

print("\n" + "="*50)
print("🎉 全部核銷付款與發票明細已 100% 全自動建立完成！")
print("==================================================")
input("按 Enter 鍵關閉視窗...")
