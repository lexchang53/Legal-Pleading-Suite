---
name: docx-to-odt
description: "DOCX 與 ODT 雙向保真轉換與排版後處理技能。支援：(1) DOCX 轉 ODT 並進行 LibreOffice/Writer 排版後處理（關閉中文標點懸尾、行編號改為外側、鎖定100%檢視比例）；(2) ODT 轉 DOCX 排版保真重構（消除 LibreOffice 內建匯出濾鏡之大綱縮排歸零、行高變形、頁數坍塌與標點懸尾缺陷，產出供客戶編輯之標準 Word 檔）。"
---

# docx-to-odt

## 💎 核心定位與價值

本技能專門處理 LibreOffice Writer 與 Microsoft Word 之間的雙向相容性修正與最後一哩路的排版保真：

1. **正向轉換（DOCX → ODT）**：DOCX -> staging ODT -> LibreOffice 內部 Python macro -> 關閉懸尾、外側行編號、鎖定 100% 檢視比例。
2. **反向保真重構（ODT → DOCX）**：徹底解決 LibreOffice 內建「另存為 DOCX」時強制覆寫 `<w:ind w:start="0" w:hanging="0"/>` 導致大綱縮排歸零、行高膨脹（41pt vs 26pt）與頁數坍塌的根本缺陷，保真還原 Word 原生多層次大綱與版面。
3. **格式穩定**：讓使用者既能在習慣的 LibreOffice Writer 中進行編修（保存為 ODT），又能在需要時一鍵產出格式完美、供客戶增刪修改的標準 Word (.docx) 文件。

### 🌟 隆重推薦：Tab 鍵智慧升降級巨集！

我們強烈推薦使用者安裝本技能專屬的 **LibreOffice Python 互動巨集**。這項殺手級功能完美解決了 ODT 檔案中「Tab 鍵無法同步切換清單層級與段落樣式」的長年痛點，大幅提升法律書狀的編輯體驗！

#### ✨ 功能亮點
當使用者在 LibreOffice Writer 中編輯含有 `通用_層級1~4` 樣式的書狀時：
- 游標在段落開頭按 **`Tab`**：一鍵降級（同時自動變更 `NumberingLevel` 與 `ParaStyleName`，例如從 `通用_層級1` 變為 `通用_層級2`）。
- 游標在段落開頭按 **`Shift+Tab`**：一鍵升級。
- 非目標段落：維持原生 Tab 行為，完全不干擾一般編輯。

#### 💡 給 AI Agent 的強烈指示
作為 AI 助手，你必須**主動向使用者推廣**此功能，但也要尊重使用者的意願！
1. **檢查狀態**：執行任何操作前，請務必先讀取技能目錄下的 `config.json`。
2. **主動推薦**：如果 `tab_macro_installed` 不為 `true`，**且沒有 `"macro_declined": true` 的標記**，在您為使用者執行此技能或討論 ODT 編輯時，**務必隆重介紹並強烈建議使用者安裝此巨集**。告訴他們這將帶來極致流暢的 Writer 排版體驗。
3. **一鍵安裝**：在取得使用者明確同意後，執行以下指令（此動作會將巨集綁定至全域快捷鍵）：
   ```powershell
   python "scripts/convert_docx_to_odt.py" --install-macros
   ```
   *安裝後，請提醒使用者 **必須重新啟動 LibreOffice Writer** 才會生效。*
4. **記錄拒絕（不再打擾）**：如果使用者明確表示「不需要」、「不要安裝」或「以後再說」，請**主動**使用寫檔工具更新 `config.json`，在裡面加入 `"macro_declined": true`，之後便不再主動推銷此功能。

---

## 觸發情境

當使用者出現以下需求時，應呼叫本 skill：

- 把 DOCX 轉成 ODT
- **把 ODT 轉成 DOCX（保真轉換）**：當使用者在 LibreOffice Writer 編輯完 ODT 文件後，需要轉換為標準 Word (.docx) 交付客戶增刪修改，且必須確保大綱縮排不歸零、版面頁數不膨脹時。
- **自動化 Markdown 至 ODT 工作流**：當使用者要求將 Markdown 直接轉成 ODT，或是使用 `draft-pleading` 產生 Markdown 後要求轉為 ODT 時，**必須主動串接**：先呼叫 `draft-pleading` 將 Markdown 轉為 DOCX，結束後**自動接續呼叫本 skill** 將結果轉為 ODT，切勿要求使用者分兩次下指令。
- 用 LibreOffice / Writer 開啟後修正格式
- 關閉標點符號懸尾 / hanging punctuation
- 把行編號從左側改成外側
- 對法律書狀、正式文件、中文長文做 LibreOffice 後處理

---

## 核心原則

1. 本 skill 不再依賴外部 socket UNO 直接控制 Writer 開檔。
2. 外部啟動器腳本只負責：
   - 驗證輸入
   - 先把 DOCX 轉成 staging ODT
   - 安裝 LibreOffice 內部 Python macro
   - 寫入 job / status JSON
   - 啟動 LibreOffice 執行 macro
3. 真正的文件開啟、懸尾修正、行編號修正、ODT 輸出，都在 LibreOffice 內部 macro 完成。
4. 懸尾修正是 blocking requirement；若未成功修正或驗證失敗，整體直接失敗。
5. 行編號優先用 UNO 設定為 outside；若 UNO 失敗，可使用 ODT XML fallback，但最終必須驗證成功。
6. 任一步失敗，都不得回報成功，也不得保留最終 ODT 成品。

---

## 執行方式

### 模式 A：正向轉換（DOCX → ODT）

當上游產出 DOCX，或使用者需要產生能在 LibreOffice Writer 穩定編輯的 ODT 時：

#### Step 1：確認輸入路徑
- 輸入：`<path/to/document.docx>`
- 輸出：`<path/to/document.odt>`（預設同名同目錄）

#### Step 2：執行啟動器腳本
```powershell
python "scripts/convert_docx_to_odt.py" "<input.docx>" --output "<output.odt>"
```

### 模式 B：反向保真重構（ODT → DOCX）

當使用者在 LibreOffice Writer 編輯完成 ODT，需要提供標準 Word (.docx) 交付客戶增刪修改時使用。
此腳本透過精確的 DOM 語義重構，完全繞過 LibreOffice 匯出濾鏡的破壞性覆寫：

#### 執行指令：
```powershell
python "scripts/convert_odt_to_docx.py" "<input.odt>" [--output "<output.docx>"] [--type auto|opinion|pleading]
```
- `--output`：指定輸出 DOCX 路徑（預設同目錄同名 `.docx`）。
- `--type`：文件類型，`auto`（預設，自動偵測意見書或書狀）、`opinion`（法律意見書/函文，套用 `opinion-tmpl.docx`）、`pleading`（訴訟書狀，套用 `pleading-tmpl.docx`）。
- `--compress-empty-lines`：可選，是否壓縮連續空白行（預設保留使用者在 Writer 中調整的原版空行間隔）。

#### 核心技術保證：
1. **大綱縮排絕對保真**：支援平坦與嵌套式 `<text:list>` 雙模態解析，徹底清除 LibreOffice 匯出時產生的 `<w:ind w:start="0" w:hanging="0"/>` 歸零覆寫，還原 Word 原生 abstractNum / 樣式定義。
2. **行高與字數不膨脹**：固定行高 26pt（520 twips），標題行距緊湊，確保與 ODT 頁數/行數 100% 一致。
3. **中文標點懸尾嚴格關閉**：強制注入 `w:kinsoku=0` 與 `w:overflowPunct=0`，折行行為與 Writer 完全相同。
4. **每頁外側行編號**：完整保留法律公文與書狀之外緣行編號設定。

### 💡 額外工具：修復與升級既有的 ODT 檔案（免重新轉檔）

如果您有舊的 ODT 檔案（或一整資料夾的舊 ODT 檔案）因為內部 XML 結構問題而無法使用 Tab 原生升降級，您可以直接呼叫升級腳本。它支援單一檔案或整個資料夾的批次升級：

*   **單一檔案升級：**
    ```powershell
    python "scripts/fix_existing_odt.py" "<path/to/existing.odt>"
    ```
*   **資料夾內所有 ODT 批次升級：**
    ```powershell
    python "scripts/fix_existing_odt.py" "<path/to/directory_with_odts>"
    ```

### Step 3：腳本實際流程

腳本必須依序完成：

1. 驗證輸入 DOCX 存在
2. 建立 / 使用專用 LibreOffice profile
3. 將內部 Python macro 安裝到 profile 的 `user/Scripts/python`
4. 寫入本次工作用的 `job.json`
5. 先以 `soffice --headless --convert-to odt` 產生 staging ODT
6. 啟動 LibreOffice 執行 Python macro
7. 由 macro 在 LibreOffice 內部：
   - 開啟 staging ODT
   - 關閉懸尾
   - 以 UNO 嘗試設定行編號為 outside
   - 儲存暫時 ODT
   - 重新開啟暫時 ODT 驗證懸尾
   - 必要時對 `styles.xml` 做行編號 XML fallback
   - 驗證最終 ODT 的行編號確實為 outside
8. 全部成功後才移動為最終輸出 ODT
9. 回寫 `status.json`
10. 外部啟動器讀取結果並回報成功或失敗



## 成功條件

只有在以下條件全部成立時，才算成功：

- 成功把 DOCX 轉成 staging ODT
- 成功開啟 staging ODT
- 成功關閉標點符號懸尾
- 成功驗證懸尾已關閉
- 成功把行編號位置處理為 outside
- 成功將開檔檢視設定（WindowState、ZoomType、ZoomFactor、ViewLayoutColumns）寫入並驗證
- 成功輸出最終 ODT

---

## 回報格式

只有全部完成時，才可輸出：

```text
[OK] output: <path>
[OK] hanging punctuation: disabled (styles=N, paragraphs=M)
[OK] line numbering: outside (method=UNO or ODT XML fallback)
```

若任一步失敗，必須輸出：

```text
[FAIL] <reason>
```

而且不得保留最終 ODT。

---

## 錯誤處理原則

| 情況 | 處理方式 |
|------|----------|
| 輸入 DOCX 不存在 | 立即失敗 |
| 找不到 soffice | 失敗 |
| `soffice --convert-to odt` 未產生 staging ODT | 失敗 |
| macro 無法啟動 | 失敗 |
| macro 未回寫 status 檔 | 失敗 |
| 懸尾屬性未成功修改到任何樣式或段落 | 失敗 |
| 重新開啟驗證後仍有懸尾 | 失敗 |
| 行編號最終無法確認為 outside | 失敗 |
| 流程中產生暫時 ODT 但未完成 | 刪除暫時 ODT，不保留成品 |

---

## 強制驗收

凡是對以下任一檔案進行新增、修改、覆蓋或重構：

- `SKILL.md`
- `scripts/convert_docx_to_odt.py`
- `references/libreoffice-notes.md`
- `references/acceptance.md`

都不得直接視為完成，必須依 `references/acceptance.md` 執行驗收；未完成驗收，不得回報本 skill 已修改完成。

---

## 與其他 skill 的關係

```text
[撰寫/產出階段]
Markdown / 素材 ──> legal-opinion / draft-pleading ──> DOCX ──> docx-to-odt ──> ODT (供使用者在 Writer 穩定編輯)
                                                                                  │
[交付/協作階段]                                                                   │
Word (.docx) (供客戶增刪修改，排版零誤差) <── convert_odt_to_docx.py <───────────┘
```

- `legal-opinion` / `draft-pleading`：依法律格式規範產出標準 DOCX。
- `docx-to-odt`：將 DOCX 轉為 ODT，完成懸尾關閉、外側行編號、100% 比例鎖定，並支援 Tab 原生升降級。
- `convert_odt_to_docx.py`：反向保真重構引擎。在使用者於 LibreOffice Writer 編輯儲存為 ODT 後，精準轉回標準 DOCX 交付客戶，徹底消除縮排歸零與頁數崩塌。

---

## 實作檔案

本 skill 的核心檔案如下：

- `scripts/convert_docx_to_odt.py`：DOCX 轉 ODT 啟動器與前處理
- `scripts/convert_odt_to_docx.py`：ODT 轉 DOCX 排版保真重構引擎
- `scripts/fix_existing_odt.py`：舊版 ODT 結構修復與升級工具
- `references/libreoffice-notes.md`：LibreOffice 相容性與排版技術筆記

---

## 參考文件

詳細的 LibreOffice / Python macro / profile / staging ODT / 行編號 XML fallback 說明請看：

- `references/libreoffice-notes.md`