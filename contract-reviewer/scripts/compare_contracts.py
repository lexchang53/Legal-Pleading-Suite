#!/usr/bin/env python3
"""
契約對比分析腳本
對比客戶契約與公司模板，生成差異報告

使用方法:
    python compare_contracts.py <customer_md> <template_md> [--output <report_file>]
"""
import json
import re
import sys
from pathlib import Path
from difflib import SequenceMatcher


def extract_clauses(md_content: str) -> dict:
    """從 markdown 中提取條款"""
    clauses = {}
    current_section = ""
    current_clause_num = ""
    current_text = []

    lines = md_content.split('\n')

    for line in lines:
        line = line.strip()

        # 檢測章節標題 (如 "第一條 定義")
        section_match = re.match(r'\*?\*?第([一二三四五六七八九十]+)條\s*(.+?)\*?\*?$', line)
        if section_match:
            # 保存上一個條款
            if current_clause_num and current_text:
                clauses[current_clause_num] = {
                    "section": current_section,
                    "text": '\n'.join(current_text)
                }

            current_section = f"第{section_match.group(1)}條 {section_match.group(2)}"
            current_clause_num = current_section
            current_text = []
            continue

        # 檢測子條款 (如 "1.1", "2.3")
        sub_match = re.match(r'^(\d+\.\d+)\s*(.+)', line)
        if sub_match:
            # 保存上一個子條款
            if current_clause_num and current_text:
                clauses[current_clause_num] = {
                    "section": current_section,
                    "text": '\n'.join(current_text)
                }

            current_clause_num = sub_match.group(1)
            current_text = [sub_match.group(2)]
            continue

        # 普通文本行
        if line and current_clause_num:
            current_text.append(line)

    # 保存最後一個條款
    if current_clause_num and current_text:
        clauses[current_clause_num] = {
            "section": current_section,
            "text": '\n'.join(current_text)
        }

    return clauses


def compare_clauses(customer_clauses: dict, template_clauses: dict) -> dict:
    """對比兩份契約的條款"""
    comparison = {
        "only_in_customer": [],    # 僅在客戶契約中
        "only_in_template": [],    # 僅在模板中
        "different": [],           # 存在差異
        "similar": [],             # 基本相同
    }

    # 找出僅在客戶契約中的條款
    for clause_num in customer_clauses:
        if clause_num not in template_clauses:
            comparison["only_in_customer"].append({
                "clause": clause_num,
                "section": customer_clauses[clause_num]["section"],
                "text": customer_clauses[clause_num]["text"][:200]
            })

    # 找出僅在模板中的條款
    for clause_num in template_clauses:
        if clause_num not in customer_clauses:
            comparison["only_in_template"].append({
                "clause": clause_num,
                "section": template_clauses[clause_num]["section"],
                "text": template_clauses[clause_num]["text"][:200]
            })

    # 對比共同條款
    common_clauses = set(customer_clauses.keys()) & set(template_clauses.keys())
    for clause_num in common_clauses:
        customer_text = customer_clauses[clause_num]["text"]
        template_text = template_clauses[clause_num]["text"]

        similarity = SequenceMatcher(None, customer_text, template_text).ratio()

        if similarity >= 0.9:
            comparison["similar"].append({
                "clause": clause_num,
                "similarity": f"{similarity*100:.1f}%"
            })
        else:
            comparison["different"].append({
                "clause": clause_num,
                "section": customer_clauses[clause_num]["section"],
                "similarity": f"{similarity*100:.1f}%",
                "customer_text": customer_text[:300],
                "template_text": template_text[:300],
            })

    return comparison


def generate_comparison_report(comparison: dict) -> str:
    """生成對比報告"""
    report = []
    report.append("=" * 60)
    report.append("契約對比分析報告")
    report.append("=" * 60)
    report.append("")

    # 統計
    report.append("【對比統計】")
    report.append(f"  僅在客戶契約中的條款: {len(comparison['only_in_customer'])} 項")
    report.append(f"  僅在公司模板中的條款: {len(comparison['only_in_template'])} 項")
    report.append(f"  存在差異的條款: {len(comparison['different'])} 項")
    report.append(f"  基本相同的條款: {len(comparison['similar'])} 項")
    report.append("")

    # 僅在客戶契約中的條款
    if comparison["only_in_customer"]:
        report.append("\n" + "=" * 20 + " 僅在客戶契約中 " + "=" * 20)
        report.append("【注意】這些條款在公司模板中沒有，需要仔細審查")
        for item in comparison["only_in_customer"]:
            report.append(f"\n  [{item['clause']}] {item['section']}")
            report.append(f"  內容: {item['text'][:100]}...")

    # 存在差異的條款
    if comparison["different"]:
        report.append("\n" + "=" * 20 + " 條款差異 " + "=" * 20)
        for item in comparison["different"]:
            report.append(f"\n  [{item['clause']}] 相似度: {item['similarity']}")
            report.append(f"  客戶版本: {item['customer_text'][:150]}...")
            report.append(f"  模板版本: {item['template_text'][:150]}...")

    # 僅在模板中的條款
    if comparison["only_in_template"]:
        report.append("\n" + "=" * 20 + " 僅在公司模板中 " + "=" * 20)
        report.append("【建議】考慮在客戶契約中增加這些條款")
        for item in comparison["only_in_template"]:
            report.append(f"\n  [{item['clause']}] {item['section']}")
            report.append(f"  內容: {item['text'][:100]}...")

    return "\n".join(report)


def main():
    if len(sys.argv) < 3:
        print("用法: python compare_contracts.py <customer_md> <template_md> [--output <report_file>] [--json]")
        print("範例: python compare_contracts.py customer.md template.md --output comparison.txt")
        sys.exit(1)

    customer_file = sys.argv[1]
    template_file = sys.argv[2]
    output_json = "--json" in sys.argv

    # 讀取文件
    with open(customer_file, 'r', encoding='utf-8') as f:
        customer_content = f.read()
    with open(template_file, 'r', encoding='utf-8') as f:
        template_content = f.read()

    # 提取條款
    print("提取客戶契約條款...")
    customer_clauses = extract_clauses(customer_content)
    print(f"  共提取 {len(customer_clauses)} 個條款")

    print("提取模板契約條款...")
    template_clauses = extract_clauses(template_content)
    print(f"  共提取 {len(template_clauses)} 個條款")

    # 對比
    print("\n對比條款...")
    comparison = compare_clauses(customer_clauses, template_clauses)

    # 生成報告
    if output_json:
        report = json.dumps(comparison, ensure_ascii=False, indent=2)
    else:
        report = generate_comparison_report(comparison)

    # 輸出
    if "--output" in sys.argv:
        output_idx = sys.argv.index("--output")
        output_file = sys.argv[output_idx + 1]
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"\n報告已保存至: {output_file}")
    else:
        print("\n" + report)


if __name__ == "__main__":
    main()
