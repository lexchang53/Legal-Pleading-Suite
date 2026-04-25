# LibreOffice / Python Macro 技術參考

## 目錄

1. 為什麼改成 macro 架構
2. 目前成功流程
3. LibreOffice Python macro 的位置
4. Python macro 的基本規則
5. 為什麼使用 status JSON
6. staging ODT 的必要性
7. 懸尾修正
8. 行編號位置
9. XML fallback 的使用界線
10. 常見問題

---

## 1. 為什麼改成 macro 架構

在此環境中，外部 socket UNO 雖然可以成功：

- 啟動 listener
- `resolver.resolve()`
- 建立 `Desktop`

但在實際 `loadComponentFromURL()` 開檔時曾出現：

- `Binary URP bridge disposed during call`
- `loadComponentFromURL(...)` 回傳 `None`

因此，本 skill 不再依賴外部 Python 經 socket 直接控制 Writer 開檔，
而改成：

1. 外部啟動器先產生 staging ODT
2. 將 Python macro 安裝到 LibreOffice 專用 profile
3. 由 LibreOffice 內部 Python macro 真正開啟 staging ODT、修正格式、輸出最終 ODT
4. 用 `status.json` 回報結果

---

## 2. 目前成功流程

目前成功版本採用以下流程：

1. 外部腳本驗證輸入 DOCX
2. 外部腳本將 macro 安裝到：
   - `<profile>/user/Scripts/python`
3. 外部腳本寫入：
   - `job.json`
   - `status.json`（由 macro 回寫）
4. 外部腳本先執行：
   - `soffice --headless --convert-to odt`
5. 產生 staging ODT
6. LibreOffice 內部 Python macro 開啟 staging ODT
7. 在內部 UNO 環境中：
   - 關閉標點符號懸尾
   - 嘗試把行編號改為 outside
   - 儲存暫時 ODT
   - 重新開啟驗證懸尾
   - 必要時做 ODT XML fallback
8. 驗證完成後，才移動為最終 ODT

---

## 3. LibreOffice Python macro 的位置

LibreOffice 說明將 Python scripts 分成：

- My Macros（使用者層）
- Application Macros（安裝層）
- Document macros（文件層）

Windows 使用者層 Python macro 預設位置是：

```text
%APPDATA%\LibreOffice\4\user\Scripts\python
```

安裝層 Python macro 位置是：

```text
{Installation}\share\Scripts\python
```

本 skill 的實作方式是使用「專用 profile」概念，因此實際安裝位置是：

```text
<profile>/user/Scripts/python
```

這樣可避免與平常使用的 LibreOffice 設定互相干擾。

---

## 4. Python macro 的基本規則

LibreOffice 的 Python macro 是 `.py` 模組中的函式。

要讓某函式可以被 LibreOffice 呼叫，模組中必須宣告：

```python
g_exportedScripts = (run_job,)
```

在 Python macro 內，可透過 `XSCRIPTCONTEXT` 取得：

- `getDocument()`
- `getDesktop()`
- `getComponentContext()`

本 skill 使用：

```python
desktop = XSCRIPTCONTEXT.getDesktop()
```

再用 `desktop.loadComponentFromURL(...)` 開啟文件。

---

## 5. 為什麼使用 status JSON

LibreOffice 官方文件指出，從 LibreOffice 介面執行 Python scripts 時，Python 的 standard output 不可用。

因此，本 skill 不依賴 `print()` 來對外回傳成功或失敗，
而是改用檔案通訊：

- 外部啟動器寫 `job.json`
- 內部 macro 回寫 `status.json`

這樣可避免把成功 / 失敗資訊綁在不穩定的 stdout 行為上。

---

## 6. staging ODT 的必要性

這個環境中最不穩的步驟是「直接用 UNO 開 DOCX」。

雖然 `soffice --convert-to odt` 可以成功產出 ODT，
但直接對 DOCX 做 `loadComponentFromURL()` 曾回傳 `None`。
因此最終成功流程改成：

- 先把 DOCX 轉成 staging ODT
- 再讓 macro 只處理 staging ODT

這樣做的好處是：

- 避開最脆弱的 DOCX 直接載入步驟
- 讓 LibreOffice 內部修正流程只面對 Writer 原生 ODT
- 讓懸尾與行編號調整更穩定

---

## 7. 懸尾修正

標點符號懸尾的核心屬性是：

- `ParaIsHangingPunctuation`

應同時處理兩個層級：

1. `ParagraphStyles`
2. 文件中的實際段落

範例：

```python
def disable_hanging_punctuation(doc):
    styles_changed = 0
    paragraphs_changed = 0

    styles = doc.StyleFamilies.getByName("ParagraphStyles")
    for name in styles.getElementNames():
        style = styles.getByName(name)
        try:
            style.ParaIsHangingPunctuation = False
            styles_changed += 1
        except Exception:
            pass

    enum = doc.Text.createEnumeration()
    while enum.hasMoreElements():
        elem = enum.nextElement()
        try:
            elem.ParaIsHangingPunctuation = False
            paragraphs_changed += 1
        except Exception:
            pass

    return styles_changed, paragraphs_changed
```

嚴格規則：

- 若 `styles_changed == 0` 且 `paragraphs_changed == 0`，視為失敗
- 重新開啟輸出的暫時 ODT 後，若仍有 `ParaIsHangingPunctuation = True`，視為失敗
- 懸尾修正必須透過 LibreOffice 內部 UNO 完成

---

## 8. 行編號位置

行編號位置優先嘗試 UNO。

常見策略：

1. 找文件的行編號設定物件
2. 若可用，確認行編號已啟用
3. 嘗試把位置設為 outside
4. 若不同 LibreOffice 版本的屬性名不同，再退回 XML fallback

常見候選屬性：

- `Position`
- `NumberPosition`

常見外側值候選：

- `3`
- `OUTSIDE`
- `outside`

目前成功案例中，UNO 可直接設定：

```text
NumberPosition = 3
```

但不論 UNO 是否成功，最終都應驗證 ODT 內行編號配置確實為 outside。

---

## 9. XML fallback 的使用界線

XML fallback 只應用在：

- 行編號位置改為 outside

不應用在：

- 標點符號懸尾

理由：

- 行編號設定可以在 ODT 的 `styles.xml` 以文件層級方式補正
- 懸尾是段落 / 樣式層級屬性，不應以粗糙 XML 補丁冒充完成

因此正確邏輯是：

- 懸尾：必須由 LibreOffice 內部 UNO 成功修正與驗證
- 行編號：UNO 優先，必要時 XML fallback，但最終必須驗證 outside

---

## 10. 常見問題

### Q1. 為什麼不直接用外部 UNO listener？

因為這個環境曾出現：
- `Binary URP bridge disposed during call`
- `loadComponentFromURL()` 回傳 `None`

所以實務上改成 LibreOffice 內部 macro 版更穩定。

### Q2. 為什麼不直接讓 macro 開 DOCX？

因為在這個環境中，直接用 UNO 開 DOCX 曾失敗；
先轉 staging ODT 後再處理，成功率更高。

### Q3. 為什麼不用 stdout 回報？

因為 LibreOffice 執行 Python macro 時，不應假設 stdout 可穩定供外部程序讀取；
所以改用 `job.json / status.json`。

### Q4. 為什麼失敗時不保留 ODT？

因為本 skill 的目標是一次輸出到位；
若懸尾或行編號條件未完成，保留 ODT 只會造成假成功。

### Q5. 為什麼這版仍保留 XML fallback？

因為行編號位置在不同 LibreOffice 版本中的 UNO 屬性不完全一致；
XML fallback 是行編號的保底方案，但不是懸尾的保底方案。