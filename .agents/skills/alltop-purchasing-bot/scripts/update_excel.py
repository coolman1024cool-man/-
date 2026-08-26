import pandas as pd

# 定義採購明細
data = {
    '品名': ['烘豆機MCU控制組材料'],
    '規格': ['1. MCU控制組 (esp32)\n2. 觸碰螢幕套件 (WaveShare 樹莓派 7吋電容式觸碰螢幕)\n3. HDMI 線組 (1.5 m)\n4. 鏡頭模組套件 (圓剛 PW310P 1080p)'],
    '用途及說明': ['智能咖啡烘豆機'],
    '計量單位': ['式'],
    '數量': [1],
    '單價': [102417]
}

df = pd.DataFrame(data)

# 覆寫原本的 RPA 專用 Excel
file_path = 'D:/01邱正彥1130801-/01經費/ALLTOP_採購自動化工具/採購自動化專用表.xlsx'
df.to_excel(file_path, index=False)
print('✅ Excel 已成功更新！')
