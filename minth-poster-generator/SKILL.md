---
name: minth-poster-generator
description: >-
  敏實科技大學專屬海報自動生成技能。支援「USR計畫/成果海報」與「主題講座/演講海報」兩種既定格式。
  左上角固定嵌入敏實科技大學高解析 Logo，右上角支援協辦單位 Logo（無則自動省略對齊），
  中央區域自動繪製對應之講座主題/演講人/時間地點或四大成果卡片，
  海報底部自動生成規範之行政資訊欄（指導單位、協辦單位、執行單位、聯繫方式）。
  輸出檔為可在 Canva 中匯入並保持圖層可編輯性的 PowerPoint (.pptx) 海報檔。
---

# 敏實科技大學專屬海報自動生成技能 (Minth Poster Generator)

## 簡介 (Overview)
本技能用於自動化產出符合 **敏實科技大學 (Minth University of Science and Technology)** 識別規範與行政公文標準的海報排版檔案。

輸出的 `.pptx` 海報檔可 **直接匯入 Canva**（點擊 Canva 的「上傳 (Upload)」➔ 選擇 `.pptx`），匯入後海報中的所有 Logo、標題、演講人、時間地點框線與底部行政欄位均會轉為 **Canva 內部可獨立編輯的向量圖層與文字框**。

---

## 支援海報模式 (Poster Modes)

### 模式 1：主題講座 / 演講海報 (`lecture`)
- **左上角**：固定嵌入敏實科技大學 Logo（高解析度矢量標誌）。
- **右上角**：協辦單位 Logo（若有傳入 Logo 路徑則顯示，無則自動省略）。
- **中央區域**：
  - 🎤 **講座主題與副標題**（大標題標楷體粗體）
  - 👤 **主講人 / Speaker**（姓名、單位與簡介框）
  - 📅 **活動時間與地點**（顯目時間地點資訊卡）
  - 💡 **講座亮點與精彩內容**（重點特色列表）
- **海報底部 Footer**：指導單位 │ 協辦單位 │ 執行單位 ＆ 聯絡方式橫條。

---

### 模式 2：USR 計畫 / 成果展示海報 (`project`)
- **左上角**：固定嵌入敏實科技大學 Logo。
- **右上角**：協辦單位 Logo（無則省略）。
- **中央區域**（四大卡片分區）：
  - 📌 **問題意識與目標**
  - 🛠️ **執行重點與方法**
  - 🌱 **社會效益與影響**
  - 🏆 **推動成果與亮點 / SDGs 圖示與照片**
- **海報底部 Footer**：指導單位 │ 協辦單位 │ 執行單位 ＆ 聯絡方式橫條。

---

## 快速使用 (Quick Start)

### 1. 生成「主題講座演講海報」
```bash
python ~/.gemini/config/skills/minth-poster-generator/scripts/generate_poster.py \
  --mode lecture \
  --title "AI與永續發展前瞻趨勢講座" \
  --speaker "張教授 / 敏實科技大學 人工智慧學院" \
  --time "115年9月20日 (三) 10:00 - 12:00" \
  --location "敏實科技大學 大華樓5樓國際會議廳" \
  --advisor "教育部" \
  --co-organizer "新竹縣政府" \
  --organizer "敏實科技大學 USR 實踐中心" \
  --contact "邱專員 (分機 1234 / email: usr@mitust.edu.tw)" \
  --out "D:\01邱正彥1130801-\0000-antigravuty-0811\NotebookLM_Sources\AI主題講座海報.pptx"
```

### 2. 生成「USR 計畫/成果海報」
```bash
python ~/.gemini/config/skills/minth-poster-generator/scripts/generate_poster.py \
  --mode project \
  --title "計畫1 創生共學AI及ESG智慧創新平台構建與推廣" \
  --advisor "教育部" \
  --organizer "敏實科技大學 USR 實踐中心" \
  --contact "邱專員 (分機 1234)" \
  --out "D:\01邱正彥1130801-\0000-antigravuty-0811\NotebookLM_Sources\USR計畫成果海報.pptx"
```

---

## 📥 如何將產出檔匯入 Canva 二次編輯

1. 開啟 [Canva 官網 (Canva.com)](https://www.canva.com/)。
2. 點擊右上角 **「建立設計 (Create a design)」** ➔ 選擇 **「匯入檔案 (Import file)」**（或直接將 `.pptx` 拖入 Canva 首頁）。
3. 選擇本技能生成的 `.pptx` 檔案。
4. 匯入完成後開啟，海報上的敏實科大 Logo、演講人、時間地點與底部資訊均為 **Canva 原生向量文字與圖層**，可隨時替換顏色、字型或插入更多照片！
