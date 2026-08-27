# 大工作區專案筆記 (Workspace Notes)

記錄跨專案的重大事項、通用進度與下一步計畫。

## 2026-08-26 (收工)
### 已完成
- 完成大工作區 AI 系統架構重構，確立了「全域技能」與「專案技能」的分野。
- 成功將 4 項業務技能 (ALLTOP採購、USR預算、工讀生彙整、海報生成) 遷移至專案本機資料夾 .agents/skills 內。
- 將本機所有尚未追蹤的通用技能資料夾同步備份至 GitHub 雲端 (coolman1024cool-man/minth-university-ai-skills)。

### 下一步計畫
- 統整各個子資料夾的開發進度。
- 依據後續任務需求，開發或引進新的自動化工具。

### 踩坑記錄
- 需注意區分全域與專案的 AI 技能配置，避免大腦混亂，目前已順利解決。

## 2026-08-27 (收工)
### 已完成
- 建立 canvas-drawing-app 專案 (基礎 HTML5 Canvas 塗鴉白板)。
- 分析舊有海報設計風格，大幅升級 minth-poster-generator 技能，將 generate_poster.py 擴充支援 5 款 Canva 海報模板：eco (生態)、	ech (科技)、genda (日程表)、keynote (大師開講)、grid (四宮格成果)。
- 建立 Canva 與 AI 協作流程：未來可由 AI 產出結構精確的 PPTX 骨架，再匯入 Canva 裝潢美化；使用者於 Canva 微調後的截圖，也能讓 AI 讀取並直接寫回 Python 代碼。

### 下一步計畫
- 繼續實作其他未完成的工作任務。
- 若使用者提供新的 Canva 修改截圖，進一步微調並固化 5 款海報的 Python 生成代碼。

### 踩坑記錄
- （無特別重大錯誤，流程順暢）
