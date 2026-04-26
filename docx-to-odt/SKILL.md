---
name: docx-to-odt
description: "DOCX 後處理 skill：先以 LibreOffice CLI 將 .docx 轉成 staging .odt，再由 LibreOffice 內部 Python macro 開啟 staging ODT，關閉中文標點符號懸尾（ParaIsHangingPunctuation=False）、將行編號位置改為外側（outside），最後輸出最終 .odt。當使用者要求把 DOCX 轉成 ODT、要求用 LibreOffice / Writer 後處理、要求關閉 hanging punctuation / 中文標點懸尾、要求把行編號改成外側、或要對法律書狀與正式文件做 LibreOffice 修正時觸發。本 skill 不負責 Markdown 轉 DOCX，也不取代 outline-docx 或 pleading-table。"
---

# docx-to-odt

此 skill 的定位是：

**DOCX -> staging ODT -> LibreOffice 內部 Python macro -> ODT**

上游 skill 只負責產出 DOCX；本 skill 專門處理 LibreOffice / Writer 相容性修正，特別是：

- 關閉標點符號懸尾
- 將行編號改成外側
- 輸出可在 LibreOffice Writer 中穩定編修的 ODT

---

## 觸發情境

當使用者出現以下需求時，應呼叫本 skill：

- 把 DOCX 轉成 ODT
- **自動化 Markdown 至 ODT 工作流**：當使用者要求將 Markdown 直接轉成 ODT，或是使用 `draft-pleading` 產生 Markdown 後要求轉為 ODT 時，**必須主動串接**：先呼叫 `outline-docx`（或 `pleading-table`）將 Markdown 轉為 DOCX，結束後**自動接續呼叫本 skill** 將結果轉為 ODT，切勿要求使用者分兩次下指令。
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

### Step 1：確認輸入路徑

- 輸入：`<path/to/document.docx>`
- 輸出：`<path/to/document.odt>`（預設同名同目錄）

### Step 2：執行啟動器腳本

```powershell
python "scripts/convert_docx_to_odt.py" "<input.docx>" --output "<output.odt>"
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

---

## 互動巨集功能 (Tab 升降級)

本 skill 提供一組可選的 **LibreOffice Python 互動巨集**，旨在解決 ODT 檔案中「Tab 鍵無法同步切換清單層級與段落樣式」的限制。此功能預設 **不自動安裝**，需由使用者明確下達指令安裝。

### 功能說明

安裝巨集後，當使用者在 LibreOffice Writer 中編輯含有 `通用_層級1~4` 樣式的 ODT 時：
- 游標在段落開頭按 **`Tab`**：降一級（同時更改 `NumberingLevel` 與 `ParaStyleName`，如：從 `通用_層級1` 變成 `通用_層級2`）。
- 游標在段落開頭按 **`Shift+Tab`**：升一級。
- 游標在段落中間或非目標樣式段落按 Tab/Shift+Tab：維持 LibreOffice 原生行為（插入 Tab 字元或退縮排）。

### 安裝方式

使用者只需執行一次以下指令，即可將巨集安裝至 LibreOffice 的全域 My Macros 中，並自動設定好快捷鍵。安裝後，日後開啟的 **所有** 符合樣式條件的 ODT 檔案皆可使用此功能。

```powershell
python "scripts/convert_docx_to_odt.py" --install-macros
```

> **給 AI Agent 的指示**：本功能為選用工具。**執行任何操作前，請先讀取技能目錄下的 `config.json`**。
> 1. 如果 `tab_macro_installed` 為 `true`，代表使用者已安裝，**請勿再主動提及或建議安裝**。
> 2. 如果為 `false` (或檔案不存在)，且使用者「主動要求安裝巨集」或「詢問如何解決 ODT Tab 鍵升降級問題」時，你才可協助執行上述指令。執行前，**必須先向使用者說明**：「此動作會將巨集安裝至 LibreOffice 的全域使用者設定檔（My Macros），並修改 Writer 的 Tab/Shift+Tab 快捷鍵綁定。」取得使用者明確同意後，再執行安裝。

**注意**：
1. 安裝後必須 **重新啟動 LibreOffice Writer** 才會生效。
2. 由於此動作會修改 LibreOffice 的全域使用者設定檔（綁定快捷鍵），因此預設為不主動執行，需使用者同意。

---

## 成功條件

只有在以下條件全部成立時，才算成功：

- 成功把 DOCX 轉成 staging ODT
- 成功開啟 staging ODT
- 成功關閉標點符號懸尾
- 成功驗證懸尾已關閉
- 成功把行編號位置處理為 outside
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
Markdown --> outline-docx / pleading-table --> DOCX --> docx-to-odt --> ODT
```

- `outline-docx`：產出 DOCX
- `pleading-table`：產出 DOCX
- `docx-to-odt`：LibreOffice / Writer 後處理並輸出 ODT
- **自動串接**：當使用者指令最終目標為 ODT 且輸入為 Markdown 時，應先交由前置排版技能將 Markdown 轉為 DOCX，再自動接續執行本技能。

---

## 實作檔案

本 skill 的核心檔案如下：

- `scripts/convert_docx_to_odt.py`
- `references/libreoffice-notes.md`

---

## 參考文件

詳細的 LibreOffice / Python macro / profile / staging ODT / 行編號 XML fallback 說明請看：

- `references/libreoffice-notes.md`