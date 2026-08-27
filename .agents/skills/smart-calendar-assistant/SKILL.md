---
name: smart-calendar-assistant
description: 智慧行事曆小幫手。當使用者上傳對話截圖、會議紀錄或用文字要求新增行程時載入。自動判讀內容、分類，並呼叫 API 寫入 Google Calendar。
---

# 智慧行事曆小幫手 (Smart Calendar Assistant)

## 技能目標
自動從使用者的自然語言對話、圖片（LINE 截圖等）中萃取行程資訊（人事時地物），判斷行程分類，並呼叫 Google Calendar API 將行程自動寫入對應的日曆中。

## 行程分類定義 (Category Routing)
- **私人** (`coolman1024cool@gmail.com`): 預設日曆。個人私事、與工作或家庭無關的行程（邱正彥）。
- **USR** (`2ea36c7b744d4115471af0903de00a63f8190b7e9ac7fc900d02e96074d10b70@group.calendar.google.com`): 辦公室 USR 計畫相關、會議、長官裁示、採購、公務行程。
- **家庭** (`bbd711188c250419e86d13b1eca81958536dc5d65f0d5d0725acac7d32b9efbe@group.calendar.google.com`): 與家庭活動、家人相關的行程。

## 執行流程
1. **理解與萃取**: 閱讀使用者提供的文本或圖片，萃取出以下資訊：
   - 標題 (Summary)
   - 描述 (Description): 盡可能保留對話中的決議、詳細細節與備註。
   - 開始時間 (Start Time): 必須轉換為 ISO 8601 格式加時區 (例如 `2026-09-01T13:00:00+08:00`)。若未指定時長，預設為 1 小時。
   - 分類 (Category): 依據上述定義，判定為「私人」、「USR」或「家庭」。
2. **執行寫入**: 
   - 呼叫 `.agents/skills/smart-calendar-assistant/scripts/calendar_router.py` 進行寫入動作。
   - 必須傳遞正確的標題、時間與分類參數。
3. **回報結果**: 將執行成功的結果（包含 Google Calendar 的查看連結）回饋給使用者。

## 前置條件與依賴
- 專案根目錄（或目前執行目錄）下必須存在擁有足夠權限的 `credentials.json` 與 `token.json`。
