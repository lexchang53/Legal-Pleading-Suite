#!/usr/bin/env python3
"""
契約資訊填充腳本
根據用戶提供的公司資訊自動填充契約中的空白欄位

依賴 docx skill 的 Document 類
使用方法:
    PYTHONPATH=/path/to/docx/skill python fill_info.py <unpacked_dir> <info_json>
"""
import json
import sys
from pathlib import Path

from scripts.document import Document


def load_info(info_file: str) -> dict:
    """載入公司資訊（必須提供文件）"""
    if not info_file:
        raise ValueError("必須提供公司資訊文件")

    path = Path(info_file)
    if not path.is_file():
        raise FileNotFoundError(f"公司資訊文件不存在: {info_file}")

    with open(info_file, 'r', encoding='utf-8') as f:
        info = json.load(f)

    if not info:
        raise ValueError("公司資訊文件為空")

    return info


def set_paragraph_text(para, new_text: str) -> bool:
    """
    設置段落的完整文本
    保留第一個 w:r 的格式，刪除其他 w:r 元素

    注意：XML 解析器可能將 w:t 中的文本拆分為多個文本節點，
    所以需要清除所有子節點後重新設置文本。
    """
    runs = list(para.getElementsByTagName("w:r"))
    if not runs:
        return False

    # 找到第一個包含 w:t 的 run
    first_run = None
    first_t = None
    for run in runs:
        t_elems = run.getElementsByTagName("w:t")
        if t_elems:
            first_run = run
            first_t = t_elems[0]
            break

    if not first_run or not first_t:
        return False

    # 清除 w:t 中的所有子節點（可能有多個文本節點）
    while first_t.firstChild:
        first_t.removeChild(first_t.firstChild)

    # 添加新的文本節點
    owner_doc = para.ownerDocument
    text_node = owner_doc.createTextNode(new_text)
    first_t.appendChild(text_node)

    # 刪除第一個 run 中其他的 w:t 元素
    t_elems = list(first_run.getElementsByTagName("w:t"))
    for t in t_elems[1:]:
        t.parentNode.removeChild(t)

    # 刪除其他所有包含 w:t 的 run（保留不含文本的 run，如書籤等）
    for run in runs:
        if run == first_run:
            continue
        t_elems = run.getElementsByTagName("w:t")
        if t_elems:
            run.parentNode.removeChild(run)

    return True


def get_paragraph_text(para, editor) -> str:
    """獲取段落的完整文本（去除空白字元）"""
    text = editor._get_element_text(para)
    # 去除所有空白字元，只保留實際文本
    return ''.join(text.split())


def find_and_fill(doc, info: dict, fill_party: str = None) -> int:
    """
    尋找並填充資訊

    參數:
        doc: Document 對象
        info: 公司資訊字典
        fill_party: 填充哪一方的資訊，"甲方" 或 "乙方"，None 則自動檢測
    """
    filled_count = 0
    editor = doc["word/document.xml"]
    all_paras = list(editor.dom.getElementsByTagName("w:p"))

    # 記錄已填充的段落索引
    filled_indices = set()

    # 首先確定我們是甲方還是乙方
    # 透過尋找契約中哪一方的資訊是空的來判斷
    if fill_party is None:
        for i, para in enumerate(all_paras):
            text = get_paragraph_text(para, editor)
            # 如果乙方後面緊跟著住所（說明乙方資訊為空），則我方是乙方
            if text == "乙方：":
                next_text = get_paragraph_text(all_paras[i + 1], editor) if i + 1 < len(all_paras) else ""
                if next_text == "住所：":
                    fill_party = "乙方"
                    break
            elif text == "甲方：":
                next_text = get_paragraph_text(all_paras[i + 1], editor) if i + 1 < len(all_paras) else ""
                if next_text == "住所：":
                    fill_party = "甲方"
                    break

    if fill_party:
        print(f"檢測到需要填充: {fill_party}")
    else:
        print("無法自動檢測填充方，將填充所有空白欄位")

    # 狀態機：跟蹤當前在填充哪一方的資訊
    current_party = None  # "甲方" 或 "乙方" 或 None
    in_bank_section = False  # 是否在銀行帳戶資訊區域

    for i, para in enumerate(all_paras):
        if i in filled_indices:
            continue

        text = get_paragraph_text(para, editor)

        # 檢測當前段落屬於哪一方
        if text == "甲方：" or text.startswith("甲方："):
            current_party = "甲方"
        elif text == "乙方：" or text.startswith("乙方："):
            current_party = "乙方"

        # 檢測銀行帳戶資訊區域（通常在 6.3 節）
        if "收款帳戶資訊" in text or "6.3" in text:
            in_bank_section = True

        # 根據 fill_party 決定是否填充
        should_fill = (fill_party is None) or (current_party == fill_party) or in_bank_section

        if not should_fill and current_party and current_party != fill_party:
            continue

        # 填充規則
        if text == "甲方：" or text == "乙方：":
            if "company_name" in info and info["company_name"]:
                party_label = text  # "甲方：" 或 "乙方："
                new_text = f"{party_label}{info['company_name']}"
                if set_paragraph_text(para, new_text):
                    print(f"✓ 填充: {party_label} -> {info['company_name']}")
                    filled_count += 1
                    filled_indices.add(i)

        elif text == "住所：":
            if "registered_address" in info and info["registered_address"]:
                if set_paragraph_text(para, f"住所：{info['registered_address']}"):
                    print(f"✓ 填充: 住所")
                    filled_count += 1
                    filled_indices.add(i)

        elif "聯絡人" in text and "電話" in text and len(text) < 20:
            contact = info.get("contact_person", "")
            phone = info.get("contact_phone", "")
            if contact or phone:
                if set_paragraph_text(para, f"聯絡人：{contact} 電話：{phone}"):
                    print(f"✓ 填充: 聯絡人和電話")
                    filled_count += 1
                    filled_indices.add(i)

        elif text == "信箱：":
            if "contact_email" in info and info["contact_email"]:
                if set_paragraph_text(para, f"信箱：{info['contact_email']}"):
                    print(f"✓ 填充: 信箱")
                    filled_count += 1
                    filled_indices.add(i)

        elif text == "開戶行：":
            if "bank_name" in info and info["bank_name"]:
                if set_paragraph_text(para, f"開戶行：{info['bank_name']}"):
                    print(f"✓ 填充: 開戶行")
                    filled_count += 1
                    filled_indices.add(i)

        elif text == "開戶名：":
            if "account_name" in info and info["account_name"]:
                if set_paragraph_text(para, f"開戶名：{info['account_name']}"):
                    print(f"✓ 填充: 開戶名")
                    filled_count += 1
                    filled_indices.add(i)

        elif text == "帳號：":
            if "bank_account" in info and info["bank_account"]:
                if set_paragraph_text(para, f"帳號：{info['bank_account']}"):
                    print(f"✓ 填充: 帳號")
                    filled_count += 1
                    filled_indices.add(i)

    return filled_count


def main():
    if len(sys.argv) < 3:
        print("用法: PYTHONPATH=/path/to/docx/skill python fill_info.py <unpacked_dir> <info_json> [party]")
        print("")
        print("參數:")
        print("  unpacked_dir  解壓縮後的文件目錄")
        print("  info_json     公司資訊 JSON 文件（必填）")
        print("  party         填充哪一方：甲方 或 乙方（可選，自動檢測）")
        print("")
        print("JSON 檔案格式範例:")
        print('''{
  "company_name": "XX科技有限公司",
  "legal_representative": "張三",
  "registered_address": "XX市XX區XX路XX號",
  "contact_person": "李四",
  "contact_phone": "138XXXXXXXX",
  "contact_email": "contact@example.com",
  "bank_name": "XX銀行XX支行",
  "bank_account": "XXXXXXXXXX",
  "account_name": "XX科技有限公司"
}''')
        sys.exit(1)

    unpacked_dir = sys.argv[1]
    info_file = sys.argv[2]
    fill_party = sys.argv[3] if len(sys.argv) > 3 else None

    # 驗證 party 參數
    if fill_party and fill_party not in ["甲方", "乙方"]:
        print(f"錯誤: party 參數必須是 '甲方' 或 '乙方'，而不是 '{fill_party}'")
        sys.exit(1)

    # 檢查目錄
    if not Path(unpacked_dir).is_dir():
        print(f"錯誤: 目錄不存在 - {unpacked_dir}")
        sys.exit(1)

    # 載入公司資訊
    try:
        print(f"載入公司資訊: {info_file}")
        info = load_info(info_file)
        print(f"已載入 {len(info)} 個欄位")
    except (ValueError, FileNotFoundError) as e:
        print(f"錯誤: {e}")
        sys.exit(1)

    # 初始化文件
    print(f"\n初始化文件: {unpacked_dir}")
    doc = Document(unpacked_dir)

    # 填充資訊
    print("\n開始填充資訊...")
    filled_count = find_and_fill(doc, info, fill_party)

    if filled_count == 0:
        print("\n警告: 未找到可填充的欄位")
        print("提示: 確保契約中有空白的 '欄位名：' 格式的占位符")
    else:
        print(f"\n共填充 {filled_count} 個欄位")

    # 保存文件
    doc.save()
    print("文件已保存")


if __name__ == "__main__":
    main()
