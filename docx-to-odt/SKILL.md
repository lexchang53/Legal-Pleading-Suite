---
name: docx-to-odt
description: "DOCX 轉 ODT 與排版後處理技能。當使用者要求將 DOCX 轉換為 ODT、將書狀/Markdown直接轉為 ODT（此時需配合 draft-pleading 串接）、使用 LibreOffice/Writer 進行格式修正、關閉中文標點符號懸尾 (hanging punctuation)、將行編號改為外側 (outside)、或是對法律書狀與正式文件做後處理時，務必觸發此技能。"
---

# docx-to-odt

## 💎 核心定位與價值

本技能專門處理 LibreOffice / Writer 相容性修正與最後一哩路的排版優化。
上游技能（如 `draft-pleading`）負責產出結構正確的 DOCX；本技能則接手處理：

1. **轉換流程**：DOCX -> staging ODT -> LibreOffice 內部 Python macro -> ODT
2. **懸尾關閉**：徹底關閉中文標點符號懸尾
3. **行編號修正**：將行編號改成外側（outside）
4. **格式穩定**：輸出可在 LibreOffice Writer 中穩定編修的 ODT

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
Markdown --> draft-pleading --> DOCX --> docx-to-odt --> ODT
```

- `draft-pleading`：產出 DOCX (含排版與大綱控制)
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