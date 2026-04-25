import sys, json, traceback; sys.stdout.reconfigure(encoding="utf-8"); import build_issue_table; from pathlib import Path;
try:
    payload = json.loads(Path(r"C:\Users\lex\.gemini\antigravity\brain\dc3625d1-bfee-4b01-ab13-74519cce2781\issue_payload.json").read_text(encoding="utf-8"))
    build_issue_table.build_issue_table(payload, Path(r"C:\Users\lex\.gemini\antigravity\skills\pleading-table-builder\assets\table-tmpl.docx"), Path(r"C:\Users\lex\.gemini\antigravity\brain\dc3625d1-bfee-4b01-ab13-74519cce2781\out.docx"), Path(r"c:\Users\lex\Dropbox\My Work\工作文件\文稿\台灣士瑞克公司(文稿)\林艾彤(文稿)\data\26-0209-士瑞克_林艾彤二審上訴理由狀.docx"))
except Exception as e:
    print(traceback.format_exc())
