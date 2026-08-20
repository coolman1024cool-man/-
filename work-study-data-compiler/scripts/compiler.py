# -*- coding: utf-8 -*-
import json
import pandas as pd
import os
from datetime import datetime

def build_global_dob_map(base_dir):
    dob_map = {}
    gender_map = {}
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            if f.endswith('.xlsx') and not f.startswith('~$'):
                filepath = os.path.join(root, f)
                try:
                    df = pd.read_excel(filepath)
                    # Check common column names
                    id_cols = [c for c in df.columns if isinstance(c, str) and '身分證' in c]
                    dob_cols = [c for c in df.columns if isinstance(c, str) and '出生' in c]
                    gender_cols = [c for c in df.columns if isinstance(c, str) and '性別' in c]
                    
                    if id_cols:
                        id_col = id_cols[0]
                        if dob_cols:
                            dob_col = dob_cols[0]
                            for _, row in df.dropna(subset=[id_col, dob_col]).iterrows():
                                dob_map[str(row[id_col]).strip()] = str(row[dob_col]).strip()
                        if gender_cols:
                            gender_col = gender_cols[0]
                            for _, row in df.dropna(subset=[id_col, gender_col]).iterrows():
                                gender_map[str(row[id_col]).strip()] = str(row[gender_col]).strip()
                except Exception as e:
                    print(f"Skipping {filepath} due to read error: {e}")
    return dob_map, gender_map

def main():
    # 1. Gather intermediate data 
    data_path = r'C:\Users\user\.gemini\antigravity\scratch\final_data_merged.json'
    if not os.path.exists(data_path):
        print(f"Error: {data_path} not found. Agent needs to extract PDF data to this file first.")
        return

    with open(data_path, encoding='utf-8') as f:
        data = json.load(f)

    # 2. Build global DOB and Gender maps from the entire folder
    base_info_dir = r'D:\01邱正彥1130801-\04工讀生相關'
    dob_map, gender_map = build_global_dob_map(base_info_dir)

    valid_records = []
    skipped_records = []

    for r in data:
        id_num = r.get('ID', '').strip()
        
        # Hardcode Unit
        r['單位'] = 'USR中心'

        # Check DOB
        dob = dob_map.get(id_num)
        if pd.isna(dob) or str(dob).strip() == '':
            skipped_records.append(f"{r.get('姓名', '未知')} ({id_num}) - 找不到出生年月日")
            continue
        
        try:
            r['出生年月日'] = str(int(float(dob)))
        except:
            r['出生年月日'] = str(dob)

        # Update gender
        inferred_gender = ''
        if len(id_num) == 10:
            if id_num[1] == '1': inferred_gender = '男'
            elif id_num[1] == '2': inferred_gender = '女'
        
        user_g = gender_map.get(id_num)
        if pd.notna(user_g) and str(user_g).strip() != '':
            r['性別'] = str(user_g).strip()
        else:
            r['性別'] = inferred_gender

        valid_records.append(r)

    # 4. Generate Output Excel
    out_cols = ['編號', '單位', '工讀月份', '學生班級', '科系', '學生姓名', '性別', '學號', '出生年月日', '身分證號碼', '地址', '時數', '應領金額', '勞方保費', '實領金額', '資方保費', '備註']
    df = pd.DataFrame(columns=out_cols)
    
    for i, r in enumerate(valid_records):
        df.loc[i] = {
            '編號': i+1,
            '單位': r.get('單位', ''),
            '工讀月份': r.get('月份', ''),
            '學生班級': r.get('班級', ''),
            '科系': str(r.get('班級', '')).replace('二', '').replace('三', '').replace('一', '').replace('夜', '').replace('A', ''), 
            '學生姓名': r.get('姓名', ''),
            '性別': r.get('性別', ''),
            '學號': r.get('學號', ''),
            '出生年月日': r.get('出生年月日', ''),
            '身分證號碼': r.get('ID', ''),
            '地址': r.get('地址', ''),
            '時數': r.get('時數', ''),
            '應領金額': r.get('應領金額', ''),
            '勞方保費': r.get('勞方保費', ''),
            '實領金額': r.get('實領金額', ''),
            '資方保費': r.get('資方保費', ''),
            '備註': '由原始紙本掃描檔萃取'
        }

    if not df.empty:
        df['month_num'] = df['工讀月份'].str.extract(r'(\d+)月').astype(float)
        df['year_num'] = df['工讀月份'].str.extract(r'(\d+)年').astype(float)
        df = df.sort_values(by=['year_num', 'month_num', '學生姓名']).drop(columns=['month_num', 'year_num'])
        df['編號'] = range(1, len(df) + 1)

    today_str = datetime.now().strftime("%Y%m%d")
    out_dir = r'D:\01邱正彥1130801-\04工讀生相關\工讀名冊彙整資料'
    os.makedirs(out_dir, exist_ok=True)
    
    output_path = os.path.join(out_dir, f'114學年度助學工讀金統計_已彙整_{today_str}.xlsx')
    df.to_excel(output_path, index=False)
    
    print(f"\n[成功] 資料已彙整完畢，並輸出至雲端: {output_path}")
    print(f"總共寫入 {len(df)} 筆有效紀錄。")
    print(f"實領總金額: {df['實領金額'].sum()}")
    
    if skipped_records:
        print("\n[警告] 以下紀錄因在全域掃描中仍找不到出生年月日，已自動跳過:")
        for skip in skipped_records:
            print(f" - {skip}")

if __name__ == "__main__":
    main()
