---
name: work-study-data-compiler
description: >-
  自動掃描並彙整工讀金相關資料，結合名冊與舊有紀錄自動補齊出生年月日及性別，並產出最新日期的 Excel 統計報表。
---

# 工讀金資料自動彙整工具 (Work-Study Data Compiler)

## Overview
此技能負責從指定的多個經費目錄下，搜尋 PDF 與 Excel 檔案來提取工讀生請款及簽到紀錄，並與工讀生名冊進行比對。
若在名冊中查無出生年月日，程式會自動跳過該筆紀錄，並在結尾輸出警告提示。最後，技能會產出一份包含今天日期的完整 Excel 彙整表，存放在桌面上。

## Usage
此技能包含一個 Python 腳本：`scripts/compiler.py`。
當您收到使用者的彙整請求時，請直接透過 Python 執行該腳本。

```bash
python ~/.gemini/config/skills/work-study-data-compiler/scripts/compiler.py
```

## Rate Limiting
無外部 API 呼叫，不適用。

## Common Mistakes
- **路徑錯誤**：腳本內寫死了特定的絕對路徑（如 `D:\01邱正彥1130801-\01經費` 等），如果使用者的資料夾結構有大幅改動，需更新腳本內的路徑設定。
- **PDF 格式變更**：如果會計系統輸出的 PDF 格式改變，可能需要調整腳本內的正則表達式擷取邏輯。

