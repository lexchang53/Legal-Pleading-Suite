# JSON Payload 欄位說明與驗證規則

## 完整 Payload 結構

```json
{
  "statement_text": "",
  "statement_items": [],
  "undisputed_text": "",
  "undisputed_items": [],
  "post_table_markdown": "",
  "party_status": "上訴人（即原審被告）",
  "reason_header": "答辯原因事實",
  "issues": [],
  "evidence_request": {
    "items": [
      {
        "related_issues": ["爭點一"],
        "investigation_item": "請求調查事項",
        "target": "調查對象",
        "target_address_contact": "對象地址及聯絡方式",
        "fact_to_prove": "待證事實"
      }
    ]
  },
  "evidence_list": {
    "title": "被上訴人證據清單表",
    "items": [
      {
        "evidence_date": "1061102",
        "evidence_name": "證據名稱",
        "evidence_summary": "證據簡要內容（不超過20字）",
        "fact_to_prove": "待證事實",
        "evidence_code": "原證7",
        "court_page": "原一85-87",
        "remarks": ""
      }
    ]
  }
}
```

## 根層欄位說明

| 欄位 | 類型 | 必填 | 說明 |
|------|------|------|------|
| `statement_text` | String | 條件必填 | 單項聲明模式使用；內容直接接在「一、聲明：」之後 |
| `statement_items` | Array[String] | 條件必填 | 多項聲明模式使用；每項對應一個 `通用_層級2` 段落 |
| `undisputed_text` | String | 條件必填 | 單項不爭執事項模式使用；內容直接接在「二、不爭執事項：」之後 |
| `undisputed_items` | Array[String] | 條件必填 | 多項不爭執事項模式使用；每項對應一個 `通用_層級2` 段落 |
| `party_status` | String | ✅ 必填 | 第三段標題的當事人稱謂 |
| `reason_header` | String | ✅ 必填 | 爭點整理表第二欄標題 |
| `issues` | Array | ✅ 必填 | 所有爭點陣列，依法院審理優先順序排列；若相容舊版，可能為 `factual_issues` 與 `legal_issues` |
| `post_table_markdown` | String | ❌ 選填 | 所有表格後補充論述，使用 Markdown 格式；無則傳空字串 |
| `evidence_request` | Object | ❌ 選填 | 聲請調查證據表資料；無則省略或傳 `null` |
| `evidence_list` | Object | ❌ 選填 | **高院證據清單表**資料；無則省略或傳 `null`，為空時不產出該表格 |

## 聲明欄位規則

聲明支援兩種模式，**只能擇一**：

### 模式 A：單項聲明

```json
{
  "statement_text": "原判決廢棄；廢棄部分，被上訴人在第一審之訴駁回。",
  "statement_items": []
}
```

### 模式 B：多項聲明

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

### 驗證規則

1. `statement_text` 與 `statement_items` 必須擇一使用。
2. 不得兩者同時為空。
3. 不得同時兩者都填入實質內容。
4. `statement_items` 若有值，必須為字串陣列，且每一項均為單一聲明事項文字，不得自行加入 `（一）`、`（二）` 等前綴。

## 不爭執事項欄位規則

不爭執事項支援兩種模式，**只能擇一**：

### 模式 A：單項不爭執事項

```json
{
  "undisputed_text": "如原審判決書第8頁所列「不爭執事項」各項所載。",
  "undisputed_items": []
}
```

### 模式 B：多項不爭執事項

```json
{
  "undisputed_text": "",
  "undisputed_items": [
    "兩造於某年某月某日簽訂系爭契約。",
    "被上訴人曾於某年某月某日寄發通知。",
    "某份書面確由被上訴人作成。"
  ]
}
```

### 驗證規則

1. `undisputed_text` 與 `undisputed_items` 必須擇一使用。
2. 不得兩者同時為空。
3. 不得同時兩者都填入實質內容。
4. `undisputed_items` 若有值，必須為字串陣列，且每一項均為單一不爭執事項文字，不得自行加入 `（一）`、`（二）` 等前綴。

## `reason_header` 允許值

`reason_header` **只能**是以下兩個值之一，不得自創：

| 值 | 適用當事人身分 |
|----|---------------|
| `主張原因事實` | 原告方、一審原告、上訴人（即原審原告）等 |
| `答辯原因事實` | 被告方、一審被告、上訴人（即原審被告）等 |

## `party_status` 允許值

`party_status` **只能**從以下選項中擇一，嚴禁自創：

- `上訴人（即原審原告）`
- `上訴人（即原審被告）`
- `被上訴人（即原審原告）`
- `被上訴人（即原審被告）`
- `上訴人即被上訴人（即原審原告）`
- `上訴人即被上訴人（即原審被告）`

## 爭點物件結構

每一筆 `issues[]`（或舊版的 `factual_issues[]`、`legal_issues[]`）內的元素，都必須符合下列格式：

```json
{
  "issue_number": "爭點一",
  "description": "上訴人解僱被上訴人是否符合最後手段性原則",
  "reasons": [
    "上訴人未先採取較輕微手段即進行解僱",
    "被上訴人年資達24年，即將符合自請退休要件"
  ],
  "laws": [
    "勞動基準法第11條第2款",
    "雇主應明確告知解僱事由（最高法院相關裁判要旨參照）"
  ],
  "evidences": [
    "兩造僱用契約書（一審卷21至24頁、甲證1）"
  ]
}
```

### 各欄位填寫規則

| 欄位 | 規則 |
|------|------|
| `issue_number` | 必須用中文國字排序，如 `爭點一`、`爭點二`，嚴禁阿拉伯數字 |
| `description` | 爭點核心描述，句型完整 |
| `reasons` | 必須為字串陣列；每項一個純文字字串，不含 `\n`、`•`、`1.`、`（一）` 等排版前綴 |
| `laws` | 必須為字串陣列；事實上爭點通常可填 `["事實認定，依證據判斷"]`；法律上爭點應填具體法條、實務見解或裁判要旨。若引用臺灣高等法院裁判，縮寫為「高院」；臺灣高等法院臺中分院，縮寫為「臺中高分院」 |
| `evidences` | 必須為字串陣列；應包含卷證位置與證據編號 |

## `evidence_request` 結構

僅在需要輸出「四、聲請調查證據表」時使用。  
若無此需求，可省略 `evidence_request`，或傳 `null`。

```json
{
  "evidence_request": {
    "items": [
      {
        "related_issues": ["爭點二"],
        "investigation_item": "請函查被上訴人之勞保投保資料",
        "target": "勞工保險局",
        "target_address_contact": "臺北市中正區羅斯福路1段4號",
        "fact_to_prove": "被上訴人於系爭期間是否另有工作收入"
      }
    ]
  }
}
```

### `evidence_request` 欄位說明

| 欄位 | 類型 | 必填 | 說明 |
|------|------|------|------|
| `items` | Array | ✅ | 聲請調查證據事項陣列；若為空陣列，最終 DOCX 不輸出第四段及其下表格 |

### `evidence_request.items[]` 結構

```json
{
  "related_issues": ["爭點一", "爭點二"],
  "investigation_item": "請求調查事項",
  "target": "調查對象",
  "target_address_contact": "地址及聯絡方式",
  "fact_to_prove": "待證事實"
}
```

### `evidence_request.items[]` 規則

| 欄位 | 規則 |
|------|------|
| `related_issues` | 必須為字串陣列，只填 `爭點一`、`爭點二` 等，不填完整爭點句子 |
| `investigation_item` | 調查事項內容 |
| `target` | 調查對象名稱 |
| `target_address_contact` | 對象地址及聯絡方式；可含 `\n` 供 Word 儲存格內分行 |
| `fact_to_prove` | 待證事實內容 |

## `extract_issue_json.py` 驗證規則

腳本提取 JSON 後，至少應驗證：

1. 必填根層欄位存在：`party_status`、`reason_header`、`factual_issues`、`legal_issues`。
2. `statement_text` / `statement_items` 必須符合二擇一規則。
3. `undisputed_text` / `undisputed_items` 必須符合二擇一規則。
4. `reason_header` 必須是 `主張原因事實` 或 `答辯原因事實`。
5. `factual_issues` 與 `legal_issues` 必須為陣列。
6. 每個爭點物件必須包含 `issue_number`、`description`、`reasons`、`laws`、`evidences`。
7. `evidence_request` 若存在，必須為物件；其中 `items` 必須為陣列。
8. `evidence_request.items[]` 每筆資料都必須包含 `related_issues`、`investigation_item`、`target`、`target_address_contact`、`fact_to_prove`。

---

## `evidence_list` 欄位說明

僅在需要產出高院證據清單表時使用。為空或省略時，不產出該表格。

詳細規則請參閱：`references/evidence-list-rules.md`

### `evidence_list` 根層欄位

| 欄位 | 類型 | 必填 | 說明 |
|------|------|------|------|
| `title` | String | ✅ | 格式為「[稱謂]證據清單表」，如「被上訴人證據清單表」 |
| `items` | Array | ✅ | 證據項目陣列；為空陣列時不產出表格 |

### `evidence_list.items[]` 結構

| 欄位 | 類型 | 必填 | 說明 |
|------|------|------|------|
| `evidence_date` | String | ✅ | 7 位民國紀元壓縮格式（如 `1061102`）；無法確認時填 `0000000` |
| `evidence_name` | String | ✅ | 證據名稱（一望即知的簡稱） |
| `evidence_summary` | String | ✅ | 證據簡要內容，建議不超過 20 字 |
| `fact_to_prove` | String | ✅ | 待證事實 |
| `evidence_code` | String | ✅ | **必須含「證」字**（如原證7、甲證1），或為「法院」；絕對不得為「附件」 |
| `court_page` | String | ✅ | 依法院卷宗頁碼代碼體系填寫（如 `原一85-87`、`更Ⅱ三56`） |
| `remarks` | String | ❌ 選填 | 備註意見；省略或空字串時儲存格留白 |