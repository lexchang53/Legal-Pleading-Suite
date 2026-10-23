#!/usr/bin/env python3
"""
契約批註腳本
為客戶契約中的風險條款添加 Word 批註

依賴 docx skill 的 Document 類
使用方法:
    PYTHONPATH=/path/to/docx/skill python add_comments.py <unpacked_dir> <risks_json> [author]
"""
import json
import sys
from pathlib import Path

from scripts.document import Document


def load_risks(risks_file: str) -> list[dict]:
    """載入風險分析結果"""
    with open(risks_file, 'r', encoding='utf-8') as f:
        return json.load(f)


def format_comment(risk: dict) -> str:
    """格式化批註內容"""
    level = risk.get("risk_level", "風險")
    description = risk.get("description", "")
    suggestion = risk.get("suggestion", "")

    return f"【{level}】{description}\n{suggestion}"


def add_comments_to_document(unpacked_dir: str, risks: list[dict], author: str = "契約審核") -> Document:
    """為文件添加批註"""
    doc = Document(unpacked_dir, author=author, initials="審")

    # 按行號排序，從後往前處理（避免行號變化影響）
    sorted_risks = sorted(risks, key=lambda r: r.get("line_number", 0), reverse=True)

    success_count = 0
    failed_risks = []

    for risk in sorted_risks:
        # 優先使用 context（更完整的上下文），其次使用 matched_text
        search_text = risk.get("context", "").replace("...", "")[:50]  # 取前50字元
        if not search_text:
            search_text = risk.get("matched_text", "")

        if not search_text:
            continue

        try:
            # 嘗試在段落級別尋找
            node = doc["word/document.xml"].get_node(tag="w:p", contains=search_text)

            if node:
                comment_text = format_comment(risk)
                doc.add_comment(start=node, end=node, text=comment_text)
                success_count += 1
                category = risk.get("category", "未知")
                level = risk.get("risk_level", "")
                print(f"✓ [{level}] {category}")
            else:
                failed_risks.append(risk)

        except Exception as e:
            failed_risks.append(risk)
            if "--verbose" in sys.argv:
                print(f"✗ 添加批註失敗: {search_text[:30]}... - {str(e)}")

    print(f"\n批註添加完成: 成功 {success_count} 項, 失敗 {len(failed_risks)} 項")

    if failed_risks and "--verbose" in sys.argv:
        print("\n未能添加批註的條款:")
        for risk in failed_risks:
            print(f"  - {risk.get('category')}: {risk.get('matched_text', '')[:30]}...")

    return doc


def main():
    if len(sys.argv) < 3:
        print("用法: PYTHONPATH=/path/to/docx/skill python add_comments.py <unpacked_dir> <risks_json> [author]")
        print("")
        print("參數:")
        print("  unpacked_dir  解壓縮後的文件目錄")
        print("  risks_json    風險分析 JSON 文件（由 analyze_risks.py 生成）")
        print("  author        批註作者名稱（可選，默認：契約審核）")
        print("  --verbose     顯示詳細資訊")
        print("")
        print("範例:")
        print("  PYTHONPATH=../.claude/skills/docx python add_comments.py ./unpacked risks.json")
        print("  PYTHONPATH=../.claude/skills/docx python add_comments.py ./unpacked risks.json 法務部")
        sys.exit(1)

    unpacked_dir = sys.argv[1]
    risks_file = sys.argv[2]

    # 解析作者參數（跳過 --verbose）
    author = "契約審核"
    for arg in sys.argv[3:]:
        if not arg.startswith("--"):
            author = arg
            break

    # 檢查目錄和文件
    if not Path(unpacked_dir).is_dir():
        print(f"錯誤: 目錄不存在 - {unpacked_dir}")
        sys.exit(1)

    if not Path(risks_file).is_file():
        print(f"錯誤: 文件不存在 - {risks_file}")
        sys.exit(1)

    # 載入風險數據
    print(f"載入風險分析結果: {risks_file}")
    risks = load_risks(risks_file)

    high = len([r for r in risks if r.get("risk_level") == "高風險"])
    mid = len([r for r in risks if r.get("risk_level") == "中風險"])
    low = len([r for r in risks if r.get("risk_level") == "低風險"])
    print(f"風險統計: 高風險 {high} | 中風險 {mid} | 低風險 {low}")

    # 添加批註
    print(f"\n開始添加批註（作者: {author}）...")
    doc = add_comments_to_document(unpacked_dir, risks, author)

    # 保存文件
    doc.save()
    print(f"\n文件已保存")


if __name__ == "__main__":
    main()
