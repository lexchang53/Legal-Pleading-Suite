#!/usr/bin/env python3
"""
契約風險條款分析腳本
分析客戶契約中的高風險條款，輸出風險報告

支持直接讀取 DOCX 文件（通過 docx skill）或 Markdown 文件

使用方法:
    # 直接分析 DOCX 文件（推薦）
    PYTHONPATH=/path/to/docx/skill python analyze_risks.py <contract.docx> [--json] [--output <file>]

    # 分析 Markdown 文件（相容舊方式）
    python analyze_risks.py <contract.md> [--json] [--output <file>]
"""
import json
import re
import sys
from pathlib import Path

# 高風險條款配置（包含關鍵字、風險說明和修改建議）
RISK_CONFIG = {
    "高風險": {
        "無限責任": {
            "patterns": [
                r"全部損失",
                r"任何損失",
                r"一切損失",
                r"所有損失",
                r"無上限",
                r"全額退還.*已支付.*全部費用",
                r"賠償.*遭受的全部損失",
            ],
            "description": "此條款要求承擔無限責任",
            "suggestion": "建議修改為：設置責任上限（如契約金額的100%），排除間接損失和預期利潤損失。"
        },
        "智慧財產權風險": {
            "patterns": [
                r"智慧財產權.*轉讓",
                r"著作權.*歸.*所有",
                r"放棄.*權利",
                r"全部.*歸.*方所有",
            ],
            "description": "此條款涉及智慧財產權歸屬或轉讓",
            "suggestion": "建議修改為：明確區分軟體產品與訂製開發的智慧財產權，保留核心技術的智慧財產權。"
        },
        "單方解約": {
            "patterns": [
                r"立即解除本契約",
                r"有權.*解除本契約",
                r"單方.*終止",
            ],
            "description": "此條款允許單方解約",
            "suggestion": "建議修改為：明確解約條件和程序，設置合理的整頓期（如15個工作日）。"
        },
    },
    "中風險": {
        "管轄權不利": {
            "patterns": [
                r"仲裁委員會",
                r"甲方所在地",
                r"對方所在地.*法院",
            ],
            "description": "此條款約定仲裁或對方所在地管轄",
            "suggestion": "建議修改為：約定我方所在地人民法院管轄。"
        },
        "整頓期過短": {
            "patterns": [
                r"(\d+)個工作日內.*改正",
                r"逾期達到(\d+)個工作日.*解除",
            ],
            "description": "整頓期限可能過短",
            "suggestion": "建議修改為：延長至7-15個工作日。"
        },
        "授權範圍擴大": {
            "patterns": [
                r"關聯公司",
                r"獨占.*許可",
                r"排他.*許可",
            ],
            "description": "授權範圍可能過大",
            "suggestion": "建議修改為：限定為單一法人實體，關聯公司使用需另行授權；使用普通許可而非獨占/排他許可。"
        },
        "退款條款": {
            "patterns": [
                r"退還.*未使用.*款項",
                r"未使用款項.*退還",
            ],
            "description": "退款條款可能對我方不利",
            "suggestion": "建議修改為：明確退款條件，區分一次性費用（不退）和服務費（按期退）的退款規則。"
        },
    },
    "低風險": {
        "自動續約": {
            "patterns": [
                r"自動續期",
                r"默認續約",
                r"自動延長",
            ],
            "description": "包含自動續約條款",
            "suggestion": "建議關注：確認續約條件和取消通知期限。"
        },
        "保密期限": {
            "patterns": [
                r"保密.*持續有效",
                r"保密義務.*終止後",
            ],
            "description": "保密期限可能未明確",
            "suggestion": "建議修改為：明確保密期限為契約終止後2-3年。"
        },
    },
}


def extract_text_from_docx(docx_path: str) -> list[dict]:
    """從 DOCX 文件中直接提取段落和條款（通過 docx skill）"""
    import subprocess
    import tempfile

    # 解包 DOCX
    with tempfile.TemporaryDirectory() as tmpdir:
        unpack_script = Path(__file__).parent.parent.parent / "docx" / "ooxml" / "scripts" / "unpack.py"
        subprocess.run(
            ["python3", str(unpack_script), docx_path, tmpdir],
            capture_output=True, check=True
        )

        # 使用 XMLEditor 讀取段落
        sys.path.insert(0, str(Path(__file__).parent.parent.parent / "docx"))
        from scripts.utilities import XMLEditor

        editor = XMLEditor(Path(tmpdir) / "word" / "document.xml")

        paragraphs = []
        current_section = ""
        current_clause = ""

        for i, para in enumerate(editor.dom.getElementsByTagName("w:p")):
            text = editor._get_element_text(para).strip()
            if not text:
                continue

            # 檢測章節標題（支持多種格式）
            if re.match(r'^第[一二三四五六七八九十\d]+條', text):
                current_section = text
            elif re.match(r'^[一二三四五六七八九十]+[、.]', text):
                current_section = text
            elif re.match(r'^\*{0,2}第[一二三四五六七八九十\d]+條', text):
                current_section = text.replace('**', '')

            # 檢測條款編號
            clause_match = re.match(r'^(\d+[\.\d]*)', text)
            if clause_match:
                current_clause = clause_match.group(1)

            paragraphs.append({
                "para_index": i,
                "section": current_section,
                "clause": current_clause,
                "text": text,
            })

        return paragraphs


def extract_text_from_markdown(md_content: str) -> list[dict]:
    """從 markdown 內容中提取段落和條款（相容舊方式）"""
    paragraphs = []
    lines = md_content.split('\n')
    current_section = ""
    current_clause = ""

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        # 檢測章節標題
        if line.startswith('**第') and '條' in line:
            current_section = line.replace('**', '')
        elif re.match(r'^第[一二三四五六七八九十\d]+條', line):
            current_section = line
        elif re.match(r'^\d+\.\d+', line):
            current_clause = line[:10]

        paragraphs.append({
            "line_number": i + 1,
            "section": current_section,
            "clause": current_clause,
            "text": line,
        })

    return paragraphs


def analyze_risks(paragraphs: list[dict]) -> list[dict]:
    """分析段落中的風險條款"""
    risks = []

    for para in paragraphs:
        text = para["text"]

        for risk_level, categories in RISK_CONFIG.items():
            for category, config in categories.items():
                for pattern in config["patterns"]:
                    match = re.search(pattern, text)
                    if match:
                        risk_item = {
                            "risk_level": risk_level,
                            "category": category,
                            "matched_text": match.group(0),
                            "context": text[:200] + "..." if len(text) > 200 else text,
                            "section": para["section"],
                            "clause": para["clause"],
                            "description": config["description"],
                            "suggestion": config["suggestion"],
                        }
                        # 根據來源添加位置資訊
                        if "para_index" in para:
                            risk_item["para_index"] = para["para_index"]
                        if "line_number" in para:
                            risk_item["line_number"] = para["line_number"]
                        risks.append(risk_item)
                        break  # 同一段落同一類別只記錄一次

    return risks


def generate_report(risks: list[dict], output_format: str = "text") -> str:
    """生成風險報告"""
    if output_format == "json":
        return json.dumps(risks, ensure_ascii=False, indent=2)

    # 文本格式報告
    report = []
    report.append("=" * 60)
    report.append("契約風險條款分析報告")
    report.append("=" * 60)
    report.append("")

    # 統計
    high_count = len([r for r in risks if r["risk_level"] == "高風險"])
    mid_count = len([r for r in risks if r["risk_level"] == "中風險"])
    low_count = len([r for r in risks if r["risk_level"] == "低風險"])

    report.append(f"風險統計：高風險 {high_count} 項 | 中風險 {mid_count} 項 | 低風險 {low_count} 項")
    report.append("")

    # 按風險等級分組輸出
    for level in ["高風險", "中風險", "低風險"]:
        level_risks = [r for r in risks if r["risk_level"] == level]
        if level_risks:
            report.append(f"\n{'='*20} {level} {'='*20}")
            for i, risk in enumerate(level_risks, 1):
                report.append(f"\n[{i}] {risk['category']}")
                report.append(f"    位置: {risk['section']} {risk['clause']}")
                report.append(f"    匹配: {risk['matched_text']}")
                report.append(f"    說明: {risk['description']}")
                report.append(f"    建議: {risk['suggestion']}")

    return "\n".join(report)


def main():
    if len(sys.argv) < 2:
        print("用法: python analyze_risks.py <contract_file> [--json] [--output <output_file>]")
        print("")
        print("參數:")
        print("  contract_file  契約文件（支持 .docx 或 .md 格式）")
        print("  --json         輸出 JSON 格式（用於批註腳本）")
        print("  --output FILE  保存報告到文件")
        print("")
        print("範例:")
        print("  python analyze_risks.py contract.docx                    # 直接分析 DOCX（推薦）")
        print("  python analyze_risks.py contract.docx --json --output risks.json")
        print("  python analyze_risks.py contract.md                      # 相容 Markdown")
        sys.exit(1)

    input_file = sys.argv[1]
    output_format = "json" if "--json" in sys.argv else "text"

    # 檢查文件
    input_path = Path(input_file)
    if not input_path.is_file():
        print(f"錯誤: 文件不存在 - {input_file}")
        sys.exit(1)

    # 根據文件類型選擇解析方式
    if input_path.suffix.lower() == '.docx':
        print(f"直接讀取 DOCX 文件: {input_file}")
        paragraphs = extract_text_from_docx(input_file)
    else:
        # 相容 Markdown 文件
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        paragraphs = extract_text_from_markdown(content)

    # 分析風險
    risks = analyze_risks(paragraphs)

    # 生成報告
    report = generate_report(risks, output_format)

    # 輸出
    if "--output" in sys.argv:
        output_idx = sys.argv.index("--output")
        output_file = sys.argv[output_idx + 1]
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"報告已保存至: {output_file}")
    else:
        print(report)

    # 返回統計資訊
    if output_format != "json":
        high_count = len([r for r in risks if r["risk_level"] == "高風險"])
        mid_count = len([r for r in risks if r["risk_level"] == "中風險"])
        low_count = len([r for r in risks if r["risk_level"] == "低風險"])
        print(f"\n共發現 {len(risks)} 個風險條款")


if __name__ == "__main__":
    main()
