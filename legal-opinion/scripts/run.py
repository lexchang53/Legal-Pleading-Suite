#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run.py — 法律意見書技能主執行入口

本腳本用以執行 Markdown 轉 DOCX 排版，並在使用者指定時，
自動串接並呼叫 docx-to-odt 技能之轉換腳本產出 ODT 檔案。

用法：
  python scripts/run.py <draft.md> [--template <tpl.docx>] [--output <out.docx>] [--odt]
"""

import sys
import os
import argparse
import subprocess

# 確保 stdout 使用 UTF-8
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    parser = argparse.ArgumentParser(description='法律意見書技能執行入口')
    parser.add_argument('draft', help='Markdown 草稿路徑')
    parser.add_argument('--template', help='Word 模板路徑')
    parser.add_argument('--output', help='輸出 DOCX 路徑')
    parser.add_argument('--odt', action='store_true', help='是否同時產出 ODT 檔')
    args = parser.parse_args()

    draft_path = os.path.abspath(args.draft)
    if not os.path.exists(draft_path):
        print(f"[ERROR] 找不到草稿檔案: {draft_path}", file=sys.stderr)
        sys.exit(1)

    # 1. 決定 DOCX 輸出路徑
    if args.output:
        docx_out = os.path.abspath(args.output)
    else:
        docx_out = os.path.splitext(draft_path)[0] + ".docx"

    # 2. 呼叫 build_opinion.py 進行排版
    script_dir = os.path.dirname(os.path.abspath(__file__))
    builder_script = os.path.join(script_dir, "build_opinion.py")
    
    cmd_build = [sys.executable, builder_script, draft_path, "--output", docx_out]
    if args.template:
        cmd_build.extend(["--template", os.path.abspath(args.template)])
        
    print(f"正在執行 Markdown 轉 DOCX 排版...")
    res = subprocess.run(cmd_build, capture_output=True, text=True, encoding='utf-8')
    if res.returncode != 0:
        print(f"[ERROR] DOCX 產出失敗！", file=sys.stderr)
        print(res.stderr, file=sys.stderr)
        sys.exit(res.returncode)
        
    print(res.stdout.strip())

    # 3. 串接 ODT 轉換（若指定 --odt）
    if args.odt:
        odt_out = os.path.splitext(docx_out)[0] + ".odt"
        # 尋找 docx-to-odt 技能位置
        # 基於使用者 active workspace 目錄，docx-to-odt 與 legal-opinion 應在同層目錄下
        skills_dir = os.path.dirname(os.path.dirname(os.path.abspath(script_dir)))
        odt_script = os.path.join(skills_dir, "docx-to-odt", "scripts", "convert_docx_to_odt.py")
        
        # fallback: 檢查絕對路徑
        if not os.path.exists(odt_script):
            odt_script = r"C:\Users\lex\.gemini\config\skills\docx-to-odt\scripts\convert_docx_to_odt.py"

        if not os.path.exists(odt_script):
            print(f"[ERROR] 找不到 docx-to-odt 轉換腳本: {odt_script}，無法執行 ODT 轉換！", file=sys.stderr)
            sys.exit(1)

        print(f"正在呼叫 docx-to-odt 技能轉換為 ODT...")
        cmd_odt = [sys.executable, odt_script, docx_out, "--output", odt_out]
        
        res_odt = subprocess.run(cmd_odt, capture_output=True, text=True, encoding='utf-8')
        if res_odt.returncode != 0:
            print(f"[ERROR] ODT 轉換失敗！", file=sys.stderr)
            print(res_odt.stderr, file=sys.stderr)
            sys.exit(res_odt.returncode)
            
        print(res_odt.stdout.strip())
        print(f"成功產出 ODT 檔: {odt_out}")

if __name__ == "__main__":
    main()
