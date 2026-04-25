---
name: pleading-table
description: "爭點整理狀 Word 排版技能。接收 draft-pleading 產出的爭點整理狀雙格式草稿（上半部人讀清單 + 下半部 JSON），以 python-docx + table-tmpl.docx 直接生成帶有爭點整理表與選填聲請調查證據表的最終 .docx。只有在使用者明確要求附表格、爭點整理表、或帶表格的爭點整理狀時才觸發；若使用者僅要求撰寫一般爭點整理狀而未要求附表格，則改由 outline-docx 處理。凡涉及爭點整理表格排版、附表格爭點整理狀 DOCX 生成時，必須使用本技能。此外，當使用者要求「高院證據清單」「高院證據清單表」「高院格式證據清單」或在高院案件用到帶「證據附卷位置」合併標頭的 7 欄證據清單表時，只要 payload 中含有 evidence_list 節點，主腳本便會自動同場產出獨立的 DOCX 附件。"
---

# Issue DOCX Builder

接收 `draft-pleading` 產出的爭點整理狀雙格式草稿（`.md`），提取最後一個 JSON 區塊，並以 `python-docx` + `table-tmpl.docx` 直接生成最終 `.docx`。

本技能支援三種表格輸出：

1. **爭點整理表**（必出）：事實上爭點與法律上爭點合併為單一表格。
2. **聲請調查證據表**（選出）：僅在 payload 中 `evidence_request.items` 非空時輸出，由兩個緊接的實體表格組成。
3. **高院證據清單表**（選出）：僅在 payload 中 `evidence_list.items` 非空時輸出，產出**獨立的 `.docx` 附件**，不嵌入爭點整理狀正文。

## 適用情境

> [!IMPORTANT]
> 只有下列情況才使用本技能：
>
> - 使用者明確要求：「爭點整理狀附表格」「含爭點整理表」「帶表格的爭點整理狀」。
> - 使用者明確要求：「高院證據清單」「高院證據清單表」「高院格式證據清單」。
> - 上游 `draft-pleading` 已產出**雙格式草稿**，亦即上半部可讀清單 + 下半部 JSON。
> - 目標是輸出最終 `.docx`，而非只修改草稿內容。

> [!CAUTION]
> 下列情況**不得**使用本技能：
>
> - 使用者只要求撰寫一般爭點整理狀，但**未要求附表格**。
> - 書狀類型不是爭點整理狀，也不是高院證據清單表，例如準備書、答辯狀、上訴理由狀、聲請狀等。
> - 只是要調整文字草稿，不需要產出附表格的 Word 成品。

## 相關技能評估

在本任務鏈中，可能相關的技能如下：

| 技能 | 是否適用 | 用途 |
|------|----------|------|
| `draft-pleading` | ✅ | 先產出爭點整理狀雙格式草稿，供本技能讀取 |
| `outline-docx` | 有條件適用 | 僅在使用者未要求附表格、只需一般爭點整理狀時使用 |
| `pleading-table` | ✅ 本技能 | 用於附表格爭點整理狀的最終 DOCX 生成 |

> [!IMPORTANT]
> 只要涉及**爭點整理表格排版**、**聲請調查證據表**、或**附表格爭點整理狀 DOCX 生成**，就應優先使用本技能，而不是 `outline-docx`。

## 絕對強制原則

> [!CAUTION]
> - **禁止使用 `docxtpl`**
> - **禁止使用 `pandoc`**
> - **禁止直接操作 XML 合併**（`merge_table.py`、`md_to_xml.py` 等舊流程一律廢止）
> - **必須使用 `python-docx`**
> - 最終 `.docx` 必須輸出到**使用者工作目錄**
> - **絕對不可**輸出到 `brain`、`artifacts` 或其他系統目錄

## 輸入格式

本技能的輸入是 `draft-pleading` 產出的雙格式 `.md` 草稿：

1. 上半部：供人閱讀與確認的清單草稿
2. 下半部：最後一個 ` ```json ... ``` ` 區塊，供腳本實際提取

> [!IMPORTANT]
> `extract_issue_json.py` 只提取**最後一個** JSON 區塊。
> 不得依賴 regex 或 Markdown 結構去反推上半部文字內容。

## JSON 最低要求

提取出的 payload 至少應包含以下欄位：

- `party_status`
- `reason_header`
- `factual_issues`
- `legal_issues`
- `statement_text` 或 `statement_items`
- `undisputed_text` 或 `undisputed_items`

選填欄位：

- `post_table_markdown`
- `evidence_request`

### 聲明欄位規則

聲明支援兩種模式，且只能擇一：

1. 單項模式：`statement_text`
2. 多項模式：`statement_items`

範例：

```json
{
  "statement_text": "原判決廢棄；廢棄部分，被上訴人在第一審之訴駁回。",
  "statement_items": []
}
```

```json
{
  "statement_text": "",
  "statement_items": [
    "原判決關於駁回上訴人後開第二至四項之訴部分暨訴訟費用之裁判，均廢棄。",
    "確認上訴人與被上訴人間僱傭關係存在。",
    "被上訴人應按月給付上訴人新臺幣若干元。"
  ]
}
```

### 不爭執事項欄位規則

不爭執事項也支援兩種模式，且只能擇一：

1. 單項模式：`undisputed_text`
2. 多項模式：`undisputed_items`

### 聲請調查證據表欄位規則

僅在需要輸出第四段及其下表格時，才提供 `evidence_request`：

```json
{
  "evidence_request": {
    "items": [
      {
        "related_issues": ["爭點二"],
        "investigation_item": "請求調查事項",
        "target": "調查對象",
        "target_address_contact": "地址及聯絡方式",
        "fact_to_prove": "待證事實"
      }
    ]
  }
}
```

若 `evidence_request` 不存在、為 `null`、或 `items` 為空陣列，最終 `.docx` 不得輸出：

- `四、聲請調查證據表`
- 表格A
- 表格B

## 工作流程

### 步驟 1：提取 JSON

執行：

```bash
python "<skill-dir>/scripts/extract_issue_json.py" "<draft.md>" \
  --output issue_payload.json
```

### 步驟 2：生成最終 DOCX

執行：

```bash
python "<skill-dir>/scripts/build_issue_table.py" issue_payload.json \
  --template "<skill-dir>/assets/table-tmpl.docx" \
  --output "<輸出路徑>/爭點整理狀.docx" \
  [--header-source "<既有書狀.docx>"]
```

參數說明：

- `issue_payload.json`：必填，步驟 1 產出的 payload。若此檔案包含 `evidence_list` 節點且 `items` 非空，腳本除了產生主書狀外，會**自動**在同一目錄產出名為「[title].docx」的高院證據清單表。
- `--template`：模板路徑，預設為 `assets/table-tmpl.docx`
- `--output`：輸出路徑，預設檔名為 `爭點整理狀.docx`
- `--header-source`：可選；指定既有書狀 `.docx` 以繼承狀首 / 狀尾。若未指定，腳本自動掃描輸出目錄中修改時間最新的 `.docx` 作為來源（排除輸出檔與 Word 暫存檔）

### 步驟 3：確認輸出

腳本完成後，至少應確認：

- 最終檔案已輸出到使用者工作目錄
- 事實上爭點與法律上爭點數量正確
- 是否成功繼承狀首 / 狀尾
- 若 `evidence_request.items` 為空，第四段與其表格確實未輸出
- **若包含高院證據清單要求**：確認輸出目錄下是否已多出該份獨立的 `.docx` 證據清單附件。

## 文件主體順序

最終 DOCX 主體順序如下：

1. 狀首
2. `一、聲明：` 或 `一、聲明：...`
3. `（一）...`、`（二）...`（僅在多項聲明時輸出）
4. `二、不爭執事項：` 或 `二、不爭執事項：...`
5. `（一）...`、`（二）...`（僅在多項不爭執事項時輸出）
6. `三、{party_status}爭點整理表`
7. 爭點整理表
8. `四、聲請調查證據表`（僅在 `evidence_request.items` 非空時輸出）
9. 聲請調查證據表（表格A + 表格B）
10. 表格後補充論述（若有）
11. 狀尾

## 中文 numbering 規則

> [!IMPORTANT]
> 正確做法：
>
> - 段落文字本身**不含**前綴
> - 套用 `通用_層級1` 或 `通用_層級2`
> - 再以 `numPr(numId, ilvl)` 指定編號層級
> - `numId` 應由腳本自 `outline-docx/assets/outline-base.docx` 注入中文 numbering 定義後取得

> [!CAUTION]
> 禁止事項：
>
> - 用 `書狀_預設` 硬寫 `一、聲明：`
> - 文字本身先手寫 `一、`、`（一）`，又同時套用 `通用_層級1` / `通用_層級2`
> - 依賴 `table-tmpl.docx` 原有 numbering 定義而不做注入

## 表格規格

### 藍圖表格

模板中應有 **2 個藍圖表格**，且腳本必須在清空主體前先鎖定其 `_tbl` XML：

1. `doc.tables[0]`：爭點整理表藍圖
2. `doc.tables[1]`：聲請調查證據表藍圖（僅欄位標題列 + 資料列）

> [!IMPORTANT]
> 聲請調查證據表僅剩單一實體表格（原表格B），已刪除表格A（㌀調查證據表表題列》與「提出人/日期列」）。

### 共通表格規則

- 所有儲存格都必須先刪除全部既有段落
- 再以 `cell.add_paragraph()` 全新建立每一段
- 以程式強制套用段落樣式
- `table.autofit = False`
- 資料列不得保留從表頭 deepcopy 繼承的 `cantSplit` / `tblHeader`

### 爭點整理表

- 單一表格，合併事實上爭點與法律上爭點
- 欄位固定為：`爭點`、`{reason_header}`、`法律依據`、`證據`
- 表頭列必須設 `tblHeader`（跟頁重複）
- 表頭列用 `爭點表_標題`
- 第 1 欄兩段都用 `爭點表_內容`
- 第 2 到 4 欄各項內容都用 `爭點表_清單`

清單格式規則：

- 僅 1 項：`•\t{text}`
- 2 項以上：`1.\t{text}`、`2.\t{text}`…

### 聲請調查證據表

為單一實體表格，第 1 列為欄位標題列，第 2 列起為資料列。

欄位固定為：

1. `編號`
2. `所涉爭點`
3. `調查事項`
4. `調查對象`
5. `對象地址及聯絡方式`
6. `待證事實`

規則：

- 欄標題列必須設 `tblHeader`（跟頁重複）
- 欄標題列可設 `cantSplit`
- `編號` 需於同一儲存格分成 2 行顯示：`編` / `號`
- `對象地址及聯絡方式` 需分成 2 行顯示：`對象地址` / `及聯絡方式`
- 資料列一律使用 `爭點表_內容`
- 若同時涉及多個爭點，`所涉爭點` 於同一儲存格分行

### 表格插入順序

聲請調查證據表為單一表格，直接插入於 anchor 前。

最終順序：

- `四、聲請調查證據表`
- 聲請調查證據表（欄位標題列 + 資料列）

## 狀尾規格

- 從「謹狀」開始的狀尾段落，均須加上 `<w:keepLines/>` + `<w:keepNext/>`
- 「謹狀」統一正規化為 `謹狀`
- 法院名稱與 `公鑒` 之間保留一個全形空白
- 法院名稱 + 公鑒 一律套用 `書狀_預設`

## 參考檔案

| 檔案 | 用途 |
|------|------|
| `assets/table-tmpl.docx` | 爭點整理狀模板 |
| `scripts/extract_issue_json.py` | 提取並驗證 JSON payload |
| `scripts/build_issue_table.py` | 生成最終 DOCX |
| `scripts/table_utils.py` | 建立爭點整理表與聲請調查證據表 |
| `references/json-schema.md` | JSON payload 規格 |
| `references/template-rules.md` | 樣式、段落與表格格式規則 |