---
name: outline-docx
description: "通用多層次大綱 DOCX 排版技能。將 Markdown 大綱草稿排版為具備四層自動編號與懸掛縮排的 Word 文件。使用 python-docx 與預設基底模板 assets/outline-base.docx，支援錨點反查、段落級 numPr 覆寫與動態編號重設。"
---

# Outline DOCX

將 Markdown 草稿排版為 `.docx`，並以 `assets/outline-base.docx` 作為預設基底模板，產出帶有四層自動編號（一、 (一) 1. (1)）與懸掛縮排的 Word 檔。

本技能是單一職責的排版引擎，專注於 Markdown → DOCX 的結構映射、模板套用與多層次編號控制；不負責特定領域（如法律、醫學）的內容生成或專用表格模型。

## 觸發條件

- 使用者要求將 Markdown 轉為 `.docx`
- 使用者要求將多層次大綱排版為正式 Word 文件
- 其他任務鏈已產出穩定 Markdown，需進行最終格式化排版

## 不適用情境

- 需要處理複雜表格資料或特定領域數據模型（應改用專門的表格技能）
- 任務主要需求是內容研發，而非格式排版
- 需要進行大規模的底層 XML 手動修改

## 絕對強制原則

> [!CAUTION]
> - **禁止使用 `pandoc`**：必須維持對 Word 樣式的精確控制
> - **必須使用 `python-docx`**：透過 `scripts/build_outline_docx.py` 執行
> - **模板依賴**：預設使用 `assets/outline-base.docx`。若需特殊樣式，應透過 `--template` 參數傳入外部模板。
> - **內容純粹性**：不得將 Markdown 語法標記（如 `#`、`1.`）直接殘留在 Word 正文中，必須轉化為對應的樣式與編號屬性。

## 核心排版機制

1. **錨點反查**：從模板中定位 `通用_層級1` 樣式的段落，反查對應的 `numId` 與 `abstractNumId`。
2. **段落級覆寫**：在 `pPr` 中明確設定 `numPr(numId, ilvl)`，確保編號層級不受全局設定縮進干擾。
3. **動態重新起算**：當出現新的第一層標題（一、）時，自動重設子層級編號。
4. **懸掛縮排與論述對齊**：自動將標題下方的純文字段落縮進至標題文字的起始點。

## 工作流程

### 步驟 1：前置準備與樣式映射（強制）

在正式執行排版腳本之前，AI **必須**：
1. **讀取語法與樣式映射表**：強制查閱 `references/md-syntax.md` 與 `references/style-mapping.md`，確認當前 Markdown 草稿中的標記（如 `#`、`1.`、`-`）應如何正確映射至 Word 的樣式層級，並於必要時先自動清理不符格式的雜用標記。
2. **自訂模板檢查（若有）**：若使用者提供了外部的自訂模板（並要求透過 `--template` 引數傳入），AI 應先檢查該模板是否具備相容的「通用多層清單」及「通用_層級」樣式設定，避免直接執行產生例外錯誤。

### 步驟 2：執行排版腳本

```bash
python "<skill-dir>/scripts/build_outline_docx.py" \
  "<draft.md路徑>" \
  --template "<skill-dir>/assets/outline-base.docx" \
  --output "<輸出docx路徑>"
```

### 步驟 3：驗證產出

執行完畢後，至少應確認：
1. 最終文件成功產出於使用者當前工作目錄。
2. 編號層級正確套用、頁碼與行編號（Line Numbering）已啟用。
3. 右側標點符號符合規則（懸尾已受控制）。

## 樣式模型

本技能僅保證支援下列核心樣式，其餘擴充樣式依模板而定：

- **通用多層清單**：所有編號段落的抽象父樣式。
- **通用_層級1**：對應 `一、`
- **通用_層級2**：對應 `(一)`
- **通用_層級3**：對應 `1.`
- **通用_層級4**：對應 `(1)`
- **Normal**：預設正文樣式。

## 排版自動化

腳本會自動執行以下排版優化：
1. **啟用全文件行編號**。
2. **停用標點符號懸尾**：確保右邊界整齊。

---

## 資源參考

| 檔案 | 用途 |
|------|------|
| `assets/outline-base.docx` | 通用排版基底模板 |
| `scripts/build_outline_docx.py` | 核心排版腳本 |
| `references/style-mapping.md` | 樣式 ID 與編號對照表 |
| `references/md-syntax.md` | Markdown 映射規則 |