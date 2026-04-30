# 法律書狀 Markdown 編輯環境設定指南

為了在 Antigravity / Cursor / VS Code 等進階編輯器中獲得最純粹、無干擾的 Markdown 撰寫體驗（避免 AI 自動改字、避免煩人的灰色預測字、以及避免編輯器自作主張的排版與括號補全），請在其他電腦上設定環境時，嚴格遵循以下兩個步驟：

## 步驟一：關閉編輯器底層 AI 預測功能（圖形介面操作）

由於 Antigravity / Cursor 內建了強大的底層 AI 補全與編輯功能（它會繞過普通的設定檔，直接在您打字時出現灰色預測字或紅色/綠色的差異比對），必須從面板手動關閉：

1. 點擊編輯器 **右下角狀態列** 的 **`Antigravity - Setting`**，進入設定面板）。
2. 找到 **Tab** 區塊，將 **`Suggestions in Editor`** (編輯器內的建議) 設為 **`Off`**。
3. （選用）找到 **Agent** 區塊，將 **`Agent Auto-Fix Lints`** 設為 **`Off`**，防止它自動幫您修正語法。

> **效果**：徹底消除打字時自動出現的「灰色預測文字」以及「紅色/綠色刪減背景」。

---

## 步驟二：修改 `settings.json` 關閉自動格式化與提示

請開啟編輯器的設定檔 `settings.json`（按下 `Ctrl + Shift + P`，輸入 `Open Settings (JSON)`），並將以下代碼加入到 JSON 物件的最外層（或直接覆蓋相關的設定）：

```json
    // --- 法律書狀專用：關閉編輯器自動補全與干擾功能 ---
    
    // 1. 關閉各種括號、引號的自動補齊與刪除
    "editor.autoClosingBrackets": "never",
    "editor.autoClosingQuotes": "never",
    "editor.autoClosingDelete": "never",
    "editor.autoClosingOvertype": "never",
    "editor.autoSurround": "never",
    
    // 2. 關閉存檔、打字、貼上時的自動格式化（避免破壞排版）
    "editor.formatOnSave": false,
    "editor.formatOnType": false,
    "editor.formatOnPaste": false,
    "editor.autoIndent": "none",
    
    // 3. 關閉打字時的下拉選字選單與自動完成提示
    "editor.quickSuggestions": {
        "other": "off",
        "comments": "off",
        "strings": "off"
    },
    "editor.suggestOnTriggerCharacters": false,
    "editor.acceptSuggestionOnEnter": "off",

    // 4. 針對 Markdown 檔案的特定強化設定
    "[markdown]": {
        "editor.quickSuggestions": {
            "other": "off",
            "comments": "off",
            "strings": "off"
        },
        "editor.acceptSuggestionOnEnter": "off",
        "editor.suggestOnTriggerCharacters": false
    }
```

> **效果**：
> * 不會再因為打出一個括號，編輯器就自動生出另一半括號，導致後續修改麻煩。
> * 不會因為按下 Enter 鍵而誤把下拉選單裡的英文單字輸入進去。
> * 存檔時，不會讓編輯器自作聰明地更改您的空白縮排與段落間距。
