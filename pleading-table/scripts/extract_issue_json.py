#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_issue_json.py
從爭點整理狀雙格式草稿（.md）提取最後一個 JSON 區塊，並進行完整驗證。

設計原則：
1. 只提取 Markdown 最後一個 ```json ... ``` 區塊
2. 驗證聲明與不爭執事項的雙模式（text/items 二擇一）規則
3. 檢查所有必填欄位與結構
4. 輸出驗證通過的標準化 JSON 到指定檔案
"""

import argparse
import json
import re
import sys
from pathlib import Path


def extract_last_json_block(md_content: str) -> dict | None:
    """提取 Markdown 最後一個 ```json ... ``` 區塊。"""
    json_blocks = re.findall(r'```json\s*\n(.*?)\n```', md_content, re.DOTALL)
    if not json_blocks:
        return None
    try:
        return json.loads(json_blocks[-1])
    except json.JSONDecodeError:
        return None


def _get_string_list(payload: dict, key: str) -> list[str]:
    """正規化字串陣列，移除空值。"""
    value = payload.get(key, [])
    if value is None:
        return []
    if not isinstance(value, list):
        raise ValueError(f"payload['{key}'] 必須為陣列")
    normalized = []
    for idx, item in enumerate(value):
        if not isinstance(item, str):
            raise ValueError(f"payload['{key}'][{idx}] 必須為字串")
        text = item.strip()
        if text:
            normalized.append(text)
    return normalized


def _validate_dual_mode_text(payload: dict, text_key: str, items_key: str, section_name: str) -> None:
    """驗證 text/items 二擇一規則。"""
    text = payload.get(text_key, "")
    items = _get_string_list(payload, items_key)

    if text is None:
        text = ""
    if not isinstance(text, str):
        raise ValueError(f"payload['{text_key}'] 必須為字串")

    has_text = bool(text.strip())
    has_items = bool(items)

    if has_text and has_items:
        raise ValueError(f"{section_name} 只能擇一使用 {text_key} 或 {items_key}，不得同時有內容")
    if not has_text and not has_items:
        raise ValueError(f"{section_name} 必須提供 {text_key} 或 {items_key} 其中之一")


def _validate_issue_item(issue: dict, issues_key: str, idx: int) -> None:
    """驗證單一爭點物件。"""
    required_fields = ["issue_number", "description", "reasons", "laws", "evidences"]
    for field in required_fields:
        if field not in issue:
            raise ValueError(f"payload['{issues_key}'][{idx}] 缺少必要欄位：{field}")

    for field in ["reasons", "laws", "evidences"]:
        items = issue.get(field, [])
        if not isinstance(items, list):
            raise ValueError(f"payload['{issues_key}'][{idx}]['{field}'] 必須為陣列")
        for item_idx, item in enumerate(items):
            if not isinstance(item, str):
                raise ValueError(f"payload['{issues_key}'][{idx}]['{field}'][{item_idx}] 必須為字串")


def _validate_evidence_request(er: dict) -> None:
    """驗證 evidence_request 結構。"""
    # 僅 items 為必填；applicant / submit_date 已從表格A刪除，不再驗證
    if "items" not in er:
        raise ValueError("payload['evidence_request'] 缺少必要欄位：items")

    items = er.get("items", [])
    if not isinstance(items, list):
        raise ValueError("payload['evidence_request']['items'] 必須為陣列")

    for idx, item in enumerate(items):
        required_fields = ["related_issues", "investigation_item", "target", "target_address_contact", "fact_to_prove"]
        for field in required_fields:
            if field not in item:
                raise ValueError(f"payload['evidence_request']['items'][{idx}] 缺少必要欄位：{field}")


def validate_payload(payload: dict) -> dict:
    """完整驗證 payload 並回傳正規化版本。"""
    # 根層必填欄位
    required_root_keys = ["party_status", "reason_header"]
    for key in required_root_keys:
        if key not in payload:
            raise ValueError(f"payload 缺少必要欄位：{key}")

    if "issues" not in payload and not ("factual_issues" in payload and "legal_issues" in payload):
        raise ValueError("payload 必須包含 'issues'，或同時包含 'factual_issues' 與 'legal_issues'（相容舊版）")

    # 聲明雙模式驗證
    _validate_dual_mode_text(payload, "statement_text", "statement_items", "聲明")

    # 不爭執事項雙模式驗證
    _validate_dual_mode_text(payload, "undisputed_text", "undisputed_items", "不爭執事項")

    # reason_header 限制
    if payload["reason_header"] not in ("主張原因事實", "答辯原因事實"):
        raise ValueError("payload['reason_header'] 必須是 '主張原因事實' 或 '答辯原因事實'")

    # party_status 限制
    allowed_status = [
        "上訴人（即原審原告）", "上訴人（即原審被告）",
        "被上訴人（即原審原告）", "被上訴人（即原審被告）",
        "上訴人即被上訴人（即原審原告）", "上訴人即被上訴人（即原審被告）"
    ]
    if payload["party_status"] not in allowed_status:
        print(f"[警告] party_status '{payload['party_status']}' 不屬於標準值，將繼續處理但建議修正")

    # 爭點陣列驗證
    for issues_key in ("issues", "factual_issues", "legal_issues"):
        if issues_key not in payload:
            continue
        issues = payload.get(issues_key, [])
        if not isinstance(issues, list):
            raise ValueError(f"payload['{issues_key}'] 必須為陣列")
        for idx, issue in enumerate(issues):
            _validate_issue_item(issue, issues_key, idx)

    # evidence_request 驗證（選填）
    er = payload.get("evidence_request")
    if er is not None:
        if not isinstance(er, dict):
            raise ValueError("payload['evidence_request'] 必須為物件或省略")
        _validate_evidence_request(er)

    # 正規化輸出
    normalized = {
        "statement_text": payload.get("statement_text", ""),
        "statement_items": _get_string_list(payload, "statement_items"),
        "undisputed_text": payload.get("undisputed_text", ""),
        "undisputed_items": _get_string_list(payload, "undisputed_items"),
        "post_table_markdown": payload.get("post_table_markdown", ""),
        "party_status": payload["party_status"],
        "reason_header": payload["reason_header"],
        "issues": payload.get("issues", []),
        "factual_issues": payload.get("factual_issues", []),
        "legal_issues": payload.get("legal_issues", []),
        "evidence_request": payload.get("evidence_request")
    }

    print("[驗證] Payload 驗證通過")
    print(f"  聲明：{'單項' if normalized['statement_text'].strip() else '多項' if normalized['statement_items'] else '無'}")
    print(f"  不爭執事項：{'單項' if normalized['undisputed_text'].strip() else '多項' if normalized['undisputed_items'] else '無'}")
    if "issues" in payload:
        print(f"  所有爭點：{len(normalized['issues'])} 個")
    else:
        print(f"  事實上爭點：{len(normalized['factual_issues'])} 個")
        print(f"  法律上爭點：{len(normalized['legal_issues'])} 個")
    if normalized["evidence_request"]:
        print(f"  聲請調查證據：{len(normalized['evidence_request'].get('items', []))} 筆")

    return normalized


def main():
    parser = argparse.ArgumentParser(description="從爭點整理狀草稿提取並驗證 JSON payload")
    parser.add_argument("input_md", help="輸入的雙格式草稿 .md 檔案")
    parser.add_argument("--output", "-o", required=True, help="輸出 JSON payload 路徑")
    args = parser.parse_args()

    input_path = Path(args.input_md)
    if not input_path.exists():
        print(f"[錯誤] 找不到輸入檔案：{input_path}", file=sys.stderr)
        sys.exit(1)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        md_content = input_path.read_text(encoding="utf-8")
    except Exception as e:
        print(f"[錯誤] 讀取檔案失敗：{e}", file=sys.stderr)
        sys.exit(1)

    print(f"[開始] 分析 {input_path.name} ...")
    payload = extract_last_json_block(md_content)

    if payload is None:
        print("[錯誤] 草稿中找不到有效的 JSON 區塊", file=sys.stderr)
        print("檢查點：", file=sys.stderr)
        print("  1. 草稿最後是否有一個完整的 ```json ... ``` 區塊？", file=sys.stderr)
        print("  2. JSON 語法是否正確（逗號、引號、大括號配對）？", file=sys.stderr)
        sys.exit(1)

    try:
        validated_payload = validate_payload(payload)
    except ValueError as e:
        print(f"[錯誤] Payload 驗證失敗：{e}", file=sys.stderr)
        sys.exit(1)

    try:
        output_path.write_text(json.dumps(validated_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[完成] 已儲存驗證通過的 payload 到：{output_path}")
    except Exception as e:
        print(f"[錯誤] 寫入輸出檔案失敗：{e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()