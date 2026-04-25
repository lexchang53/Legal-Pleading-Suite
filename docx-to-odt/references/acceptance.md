# acceptance

本文件定義 `docx-to-odt` skill 的驗收標準。  
目的不是解釋實作原理，而是確保每次修改後，skill 仍符合以下硬性要求：

- 必須成功輸出 ODT
- 標點符號懸尾必須關閉
- 行編號必須在外側
- 任一步失敗都不得假成功
- 失敗時不得保留最終 ODT 成品

---

## 1. 驗收範圍

本驗收適用於以下檔案：

- `SKILL.md`
- `scripts/convert_docx_to_odt.py`
- `references/libreoffice-notes.md`
- `references/acceptance.md`

---

## 2. 最低通過條件

每次修改後，至少必須完成以下 2 類驗收：

1. 成功案例驗收
2. 失敗案例驗收

而且只有在以下條件全部成立時，才算通過：

- 成功案例通過
- 失敗案例通過
- 最終 `.odt` 檔存在且可正常開啟
- 標點符號懸尾已關閉
- 行編號位於外側
- 文件敘述與腳本實作一致
- 沒有退化情況

---

## 3. 成功案例驗收

### 3.1 測試輸入

準備一份可實際使用的 `.docx` 測試檔，例如：

```text
C:\path\to\sample.docx
```

### 3.2 執行指令

```powershell
python "scripts/convert_docx_to_odt.py" "C:\path\to\sample.docx"
```

若需要明確指定輸出路徑，也可以使用：

```powershell
python "scripts/convert_docx_to_odt.py" "C:\path\to\sample.docx" --output "C:\path\to\sample.odt"
```

### 3.3 預期輸出

終端輸出必須包含：

```text
[OK] conversion completed
[OK] output: C:\path\to\sample.odt
[OK] hanging punctuation: disabled (...)
[OK] line numbering: outside (...)
```

若有額外細節，例如：

```text
[OK] line numbering detail: UNO 已設定 NumberPosition=3
```

可視為正常成功訊息的一部分。

### 3.4 必查項目

成功案例至少要確認：

- 最終 `.odt` 檔確實存在
- LibreOffice Writer 可正常開啟該 `.odt`
- 文件可正常再次儲存
- 標點符號懸尾已關閉
- 行編號在外側
- 輸出訊息中包含：
  - `[OK] output:`
  - `hanging punctuation: disabled`
  - `line numbering: outside`

---

## 4. 失敗案例驗收

### 4.1 案例 A：輸入檔不存在

執行：

```powershell
python "scripts/convert_docx_to_odt.py" "C:\path\to\not-found.docx"
```

### 預期結果

- 輸出 `[FAIL]`
- exit code 非 0
- 不產生最終 `.odt`

---

### 4.2 案例 B：故意製造 LibreOffice 執行失敗

可在測試副本中暫時破壞條件，例如：

- 暫時指定錯誤的 `soffice.exe` 路徑
- 或在隔離副本中故意改壞啟動條件

### 預期結果

- 輸出 `[FAIL]`
- exit code 非 0
- 不保留最終 `.odt`

---

### 4.3 案例 C：失敗後殘留檔案檢查

失敗後必須確認：

- 不存在最終輸出 `.odt`
- staging / temp / 中間檔不可被誤認為完成品
- 不得出現「雖然失敗，但成品還留著」的狀況

---

## 5. 人工排版驗收

每次重大修改後，至少人工檢查一次輸出的 `.odt`：

1. 用 LibreOffice Writer 開啟輸出檔。
2. 檢查中文段落末尾標點，確認沒有懸尾。
3. 檢查行編號位置，確認在外側。
4. 隨機抽查多個段落與樣式，確認不是只有局部被修正。
5. 再次儲存一次，確認文件未損壞。

---

## 6. 文件一致性驗收

每次修改後，需同時確認以下敘述仍與實作一致。

### 6.1 `SKILL.md`

必須仍描述為：

- `DOCX -> staging ODT -> LibreOffice 內部 Python macro -> ODT`
- 懸尾為 blocking requirement
- 行編號最終必須為 outside
- 失敗不得保留最終 ODT

### 6.2 `references/libreoffice-notes.md`

必須仍說明：

- 為何不再依賴外部 socket UNO 直接開檔
- staging ODT 的必要性
- 懸尾必須由 LibreOffice 內部 UNO 修正
- 行編號可在必要時 XML fallback
- `job.json / status.json` 的用途

### 6.3 `scripts/convert_docx_to_odt.py`

必須仍符合：

- 外部先將 DOCX 轉成 staging ODT
- 再由 LibreOffice 內部 macro 開啟 staging ODT
- 懸尾修正後有驗證步驟
- 行編號最終必須驗證為 outside
- 失敗時刪除最終 ODT

---

## 7. 退化警訊

若出現以下任一情況，視為退化，必須修正後才可接受：

- 又改回直接用外部 UNO 開 DOCX
- 懸尾未修正卻回報成功
- 行編號不是外側卻回報成功
- 失敗時仍保留最終 ODT
- `SKILL.md` 與腳本流程不一致
- `references/libreoffice-notes.md` 與腳本流程不一致
- `references/acceptance.md` 與目前實際驗收流程不一致

---

## 8. 驗收結果記錄

每次完成驗收時，建議至少記錄以下資訊：

- 測試日期
- 使用的測試 `.docx`
- 成功案例是否通過
- 失敗案例是否通過
- 是否人工確認懸尾已關閉
- 是否人工確認行編號在外側
- 是否有發現退化情況

可以簡單記在工作筆記中，不強制規定格式，但不得完全不記錄。

---

## 9. 最終判定

只有在以下條件全部成立時，才可判定這次修改合格：

- 成功案例通過
- 失敗案例通過
- 人工排版驗收通過
- 文件一致性驗收通過
- 無退化警訊

若其中任一未成立，則此次修改不得視為完成。