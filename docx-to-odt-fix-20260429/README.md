# docx-to-odt 技能修復記錄

**日期**：2026-04-29  
**環境**：本機（lexchang Windows）  
**問題描述**：`docx-to-odt` 技能無法產出 ODT 檔案，而另一台電腦可正常運作。

---

## 本資料夾內容

| 檔案 | 說明 |
|------|------|
| `README.md` | 本說明文件（診斷、討論與後續建議） |
| `diagnosis_report.md` | 詳細診斷報告（Bug 位置、原因、比較表） |
| `convert_docx_to_odt.py` | **修復後的 Python 腳本**（可直接取代技能目錄中的同名檔案） |

> **原始技能目錄不受影響**：  
> `P:\Legal-Pleading-Suite\docx-to-odt\` 中的所有檔案均未被修改。

---

## 問題根源：MACRO_SOURCE 中的 Python SyntaxError

### 位置

`scripts/convert_docx_to_odt.py` 內，嵌入於 `MACRO_SOURCE` 字串中的  
`_enforce_line_numbering_outside_xml` 函式，第 1095～1182 行（MACRO_SOURCE 字串內）。

### 錯誤類型

Python **縮排錯誤**（IndentationError / SyntaxError）：

```python
# ❌ 原始（有 Bug）的程式碼
else:
    if re.search(r"<text:linenumbering-configuration\b", text, flags=re.IGNORECASE):
    new_text = re.sub(...)        # ← 沒有縮排！應縮進 4 格
    if new_text != text:
        text = new_text
        changed = True
    else:                          # ← else 對應關係錯亂
        ...
```

```python
# ✅ 修復後的程式碼
else:
    if re.search(r"<text:linenumbering-configuration\b", text, flags=re.IGNORECASE):
        new_text = re.sub(...)    # ← 正確縮進 4 格
        if new_text != text:
            text = new_text
            changed = True
    else:
        ...
    verified = _has_outside_line_numbering(text)  # ← 縮排層次也已修正
    data = text.encode("utf-8")
```

### 失敗鏈

```
LibreOffice 載入 macro → SyntaxError → 模組無法 import
→ run_job 無法執行 → status.json 不被寫入
→ 外部啟動器等待 180 秒 timeout
→ 三個策略全部失敗
→ 報告「無法成功啟動 LibreOffice macro」
→ 無 ODT 產出
```

---

## 其他修復內容

除縮排修正外，同時調整了 XML 處理邏輯：

- **原始版本**：在 `_enforce_line_numbering_outside_xml` 中同時處理 `styles.xml` 和 `content.xml`（嘗試用 XML 直接修正懸尾），但縮排錯誤導致整段無法執行。
- **修復版本**：回歸 `.py.working` 的正確結構，`_enforce_line_numbering_outside_xml` 只處理 `styles.xml`（行編號），懸尾修正交由 LibreOffice UNO（`_disable_hanging_punctuation`）完成，符合 `libreoffice-notes.md` 的設計原則。

---

## 檔案版本比較

| 項目 | `.py.working`（原始備份） | P: 碟原始 `.py`（有 Bug） | 本資料夾 `.py`（修復後） |
|------|--------------------------|--------------------------|-------------------------|
| 大小 | 23,520 bytes | 44,859 bytes | 44,859 bytes |
| SyntaxError | ✅ 無 | ❌ 有 | ✅ 已修正 |
| `_fix_list_style_bindings` | ❌ 無 | ✅ 有 | ✅ 有 |
| `_unify_list_style_xml` | ❌ 無 | ✅ 有 | ✅ 有 |
| Tab 升降級巨集 | ❌ 無 | ✅ 有 | ✅ 有 |
| `--install-macros` 參數 | ❌ 無 | ✅ 有 | ✅ 有 |

---

## 後續處理建議

### 步驟 1：語法驗證（已完成，可重複確認）

```powershell
$py = "C:\Users\lexchang\AppData\Local\Programs\Python\Python314\python.exe"
& $py -c "import ast; src=open('convert_docx_to_odt.py', encoding='utf-8').read(); ast.parse(src); print('Syntax OK')"
```

```powershell
# 也驗證內嵌 MACRO_SOURCE 字串
& $py -c "
ns = {}
exec(open('convert_docx_to_odt.py', encoding='utf-8').read(), ns)
import ast; ast.parse(ns['MACRO_SOURCE']); print('MACRO_SOURCE Syntax OK')
"
```

### 步驟 2：實際轉換測試

```powershell
$py = "C:\Users\lexchang\AppData\Local\Programs\Python\Python314\python.exe"
& $py "convert_docx_to_odt.py" "C:\path\to\sample.docx" --output "C:\path\to\sample.odt"
```

預期輸出：
```
[OK] conversion completed
[OK] output: C:\path\to\sample.odt
[OK] hanging punctuation: disabled (styles=N, paragraphs=M)
[OK] line numbering: outside (method=...)
```

### 步驟 3：確認正常後，覆蓋技能目錄

```powershell
Copy-Item "P:\Legal-Pleading-Suite\docx-to-odt-fix-20260429\convert_docx_to_odt.py" `
    -Destination "P:\Legal-Pleading-Suite\docx-to-odt\scripts\convert_docx_to_odt.py" `
    -Force
```

> **注意**：覆蓋前請先備份原始版本，或確認 `.py.working` 仍完好。

---

## 重要補充：兩台電腦的狀況

- 雲端 P: 碟的版本（44,859 bytes）與本機**原始**版本完全相同，都有 SyntaxError。
- 若另一台電腦也從 P: 碟同步，理論上也會有同樣問題。
- 建議在另一台電腦上也執行語法驗證，確認狀況。
