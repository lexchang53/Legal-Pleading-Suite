# docx-to-odt 技能診斷報告

## 診斷結論：找到 **3 個問題**

---

## 問題 1（關鍵 Bug）：`_enforce_line_numbering_outside_xml` 縮排錯誤 → Python SyntaxError

**位置**：`scripts/convert_docx_to_odt.py`，第 1135–1157 行（MACRO_SOURCE 字串內）

**問題描述**：  
MACRO_SOURCE 中的 `_enforce_line_numbering_outside_xml` 函式，其 `if/else` 結構有縮排錯誤：

```python
# ❌ 現有（錯誤）版本
if _has_outside_line_numbering(text):
    verified = True
else:
    if re.search(r"<text:linenumbering-configuration\b", text, ...):
    new_text = re.sub(...)        # ← 縮排錯誤！應在 if 內
    if new_text != text:
        text = new_text
        changed = True
    else:                          # ← 這個 else 對應到的是外層 if，不是 re.search
        inserted = False
        ...
```

`re.search(...)` 那個 `if` 後面直接接 `new_text = re.sub(...)` 而不縮排，導致：
1. 在 **Python 語法層面是 `SyntaxError`**（`if` 區塊是空的）。
2. `else` 分支的對應關係完全錯亂。

LibreOffice 內部 Python macro 執行時，此語法錯誤會導致整個 `run_job` 函式無法載入，進而：
- macro 直接崩潰、不寫入 `status.json`
- 外部啟動器等待 timeout（180秒）後報告「無法成功啟動 macro」

> **這是導致此台電腦無法產出 ODT 的主要原因。**

**正確版本**（參照 `.py.working`）：

```python
# ✅ 正確版本
if _has_outside_line_numbering(text):
    verified = True
else:
    if re.search(r"<text:linenumbering-configuration\b", text, ...):
        new_text = re.sub(...)    # ← 正確縮排在 if 內
        if new_text != text:
            text = new_text
            changed = True
    else:
        inserted = False
        ...
    verified = _has_outside_line_numbering(text)  # ← 這行也在 else 內
    data = text.encode("utf-8")  # ← 這行在 styles.xml 分支內
```

---

## 問題 2（次要差異）：現有版本多了 `content.xml` 處理，但縮排導致 `data` 更新遺漏

**位置**：第 1121–1159 行

現有版本將 `styles.xml` 和 `content.xml` 都列入處理（第 1121 行 `if item.filename in ["styles.xml", "content.xml"]`），但 `data = text.encode("utf-8")` 這一行被放在 `if item.filename == "styles.xml"` 的外層，看起來兩種檔案都會更新。

然而，問題 1 的縮排錯誤導致這整個區塊根本無法正確執行。

---

## 問題 3（次要差異）：現有版本多了 `_fix_list_style_bindings` 和 `_unify_list_style_xml`，但這是功能增加，不是 Bug

現有 `.py`（44KB）比 `.py.working`（23KB）多了：
- `_fix_list_style_bindings(doc)` — 修復通用_層級1~4 清單樣式綁定
- `_unify_list_style_xml(odt_path)` — XML 後處理移除自動樣式的 list-style-name 覆寫
- Tab 升降級巨集（`LEVEL_MACRO_SOURCE`）和 `--install-macros` 參數

這些都是後來新增的功能，本身沒有問題，但都被問題 1 的 SyntaxError 拖累而無法執行。

---

## 修復方案

**最直接修復**：將 MACRO_SOURCE 中 `_enforce_line_numbering_outside_xml` 函式的縮排錯誤修正。

具體是第 1135–1157 行，將：
```
if re.search(...):
new_text = re.sub(...)   # ← 這行要縮進 4 格
```
改為：
```
if re.search(...):
    new_text = re.sub(...)  # ← 縮進 4 格
```

並確保 `else`、`verified = ...`、`data = text.encode("utf-8")` 的縮排層次也同步正確。

---

## 比較摘要

| 項目 | `.py.working`（正常） | `.py`（現有，有問題） |
|------|----------------------|----------------------|
| 大小 | 23,520 bytes | 44,859 bytes |
| `_fix_list_style_bindings` | ❌ 無 | ✅ 有（功能增加） |
| `_unify_list_style_xml` | ❌ 無 | ✅ 有（功能增加） |
| Tab 升降級巨集 | ❌ 無 | ✅ 有（功能增加） |
| `_enforce_line_numbering_outside_xml` 縮排 | ✅ 正確 | ❌ **SyntaxError** |
| content.xml 懸尾修正 | ❌ 僅 styles.xml | ✅ 同時處理（理論上更完整） |

