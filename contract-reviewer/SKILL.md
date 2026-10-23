---
name: contract-reviewer
description: 契約審核與對比分析技能。當用戶需要對比客戶契約與公司模板、識別高風險條款、添加批註建議、自動填充公司資訊並生成最終 DOCX 文件時，應使用此技能。適用於契約審核、條款風險分析、契約資訊填充等場景。 (project)
---

# 契約審核與對比分析

## 概述

自動化契約審核流程：對比契約差異、識別風險條款、添加批註建議、填充公司資訊。

## 使用前準備

啟動時詢問用戶提供以下文件：

```
請提供以下文件：

1. 【必填】您公司的契約模板：___
2. 【必填】需要審核的客戶契約：___
3. 【可選】公司資訊文件（用於自動填充）：___
```

## 執行流程

### 1. 風險識別（直接讀取 DOCX）

```bash
# 直接分析 DOCX 文件（推薦，無需 pandoc）
python3 "{baseDir}/scripts/analyze_risks.py" "<客戶契約.docx>"

# 生成 JSON（用於批註）
python3 "{baseDir}/scripts/analyze_risks.py" "<客戶契約.docx>" --json --output /tmp/risks.json
```

### 2. 對比分析（可選，需要 pandoc）

如需對比兩份契約的條款差異：

```bash
pandoc "<客戶契約.docx>" -o /tmp/customer.md
pandoc "<公司模板.docx>" -o /tmp/template.md
python3 "{baseDir}/scripts/compare_contracts.py" /tmp/customer.md /tmp/template.md
```

### 3. 用戶確認

展示風險分析結果，**等待用戶確認後**再添加批註。

### 4. 添加批註

```bash
python3 "{docxSkillDir}/ooxml/scripts/unpack.py" "<客戶契約.docx>" /tmp/unpacked
PYTHONPATH="{docxSkillDir}" python3 "{baseDir}/scripts/add_comments.py" /tmp/unpacked /tmp/risks.json
python3 "{docxSkillDir}/ooxml/scripts/pack.py" /tmp/unpacked "<輸出文件.docx>"
```

### 5. 資訊填充（可選）

通過 docx skill 的 Document API 直接修改段落文本：

```python
from scripts.document import Document

doc = Document("/tmp/unpacked", author="契約填充")
doc_xml = doc["word/document.xml"]

# 遍歷段落並修改
for para in doc_xml.dom.getElementsByTagName("w:p"):
    text = doc_xml._get_element_text(para)
    if text == "乙方：":
        # 修改段落中的 w:t 元素
        for t in para.getElementsByTagName("w:t"):
            if t.firstChild:
                t.firstChild.data = "乙方：台北基礎科技股份有限公司"

doc.save()
```

## 風險等級說明

| 等級 | 類型 | 說明 |
|-----|------|-----|
| 高 | 無限責任 | "全部損失"、"任何損失" |
| 高 | 智慧財產權 | "智慧財產權轉讓"、"著作權歸...所有" |
| 高 | 單方解約 | "立即解除"、"有權解除" |
| 中 | 管轄權 | "仲裁委員會"、"甲方所在地" |
| 中 | 整頓期 | "X個工作日內改正" |
| 中 | 授權範圍 | "關聯公司"、"獨占許可" |
| 低 | 自動續約 | "自動續期"、"默認續約" |

## 腳本說明

| 腳本 | 功能 | 依賴 |
|-----|------|-----|
| `analyze_risks.py` | 風險條款分析（支持 DOCX/MD） | docx skill（DOCX 模式） |
| `compare_contracts.py` | 契約對比 | pandoc |
| `add_comments.py` | 添加批註 | docx skill |

## 參考文件

- `{baseDir}/references/risk_clauses.md` - 風險條款詳細說明
- `{baseDir}/references/common_info_template.md` - 公司資訊模板
