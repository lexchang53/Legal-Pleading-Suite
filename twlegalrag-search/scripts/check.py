from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

# Windows console UTF-8 強制設定防禦
if sys.platform == "win32":
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def fail(message: str, code: int = 1) -> int:
    print(message, file=sys.stderr)
    return code


def main() -> int:
    if len(sys.argv) != 3:
        return fail("用法：python scripts/check.py <bundle.json> <answer.txt>")

    bundle_path = Path(sys.argv[1]).expanduser().resolve()
    answer_path = Path(sys.argv[2]).expanduser().resolve()

    if not bundle_path.exists():
        return fail(f"找不到 bundle 檔案：{bundle_path}")

    if not answer_path.exists():
        return fail(f"找不到回答檔案：{answer_path}")

    # 優先嘗試：直接 import 套件以 Python API 執行（可以精確獲得整體評級與控制退出碼）
    try:
        from twlegalrag.verify import citation_check
        from twlegalrag.cli import _print_report
        from twlegalrag.retrieval import Judgment

        # 載入 bundle 與 answer
        with open(bundle_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        answer = answer_path.read_text(encoding="utf-8")

        # 重建 Judgment 物件，對齊 cli.py 的轉換格式並映射 2.2.0 欄位
        hits = []
        for i, j in enumerate(data.get("judgments", [])):
            hits.append(
                Judgment(
                    rank=i + 1,
                    doc_id=j.get("doc_id", ""),
                    citation_text=j.get("citation_text", ""),
                    court_name=j.get("court_name", ""),
                    jdate=j.get("jdate", ""),
                    snippet=j.get("listing", ""),
                    citation_url=j.get("citation_url", ""),
                    citation_markdown="",
                    result_token="",
                    case_category=j.get("case_category"),
                    fulltext=j.get("fulltext_excerpt", ""),
                    cited_articles=j.get("cited_articles", []),
                    case_history=j.get("case_history"),
                    hit_excerpt=j.get("hit_excerpt"),
                    fulltext_total_chars=j.get("fulltext_total_chars"),
                    fulltext_complete=j.get("fulltext_complete", True),
                )
            )

        # 執行檢查與報表列印
        rep = citation_check(answer, hits)
        _print_report(rep)

        # 根據 overall 狀態回傳退出碼
        if rep.overall == "fail":
            return 1
        elif rep.overall == "needs_review":
            return 2  # 返回需要人工複核的中間碼
        else:
            return 0

    except ImportError:
        # 兜底方案：如果無法直接 import，則以 subprocess 呼叫 CLI 並解析 stdout 報表
        # 優先尋找 PATH 上的 twlegalrag 指令，若不在 PATH 則降級改用當前 Python 的 -m 模組模式
        exe = shutil.which("twlegalrag")
        cmd = [exe] if exe else [sys.executable, "-m", "twlegalrag"]

        try:
            result = subprocess.run(
                cmd + ["check", str(bundle_path), str(answer_path)],
                capture_output=True,
                text=True,
                encoding="utf-8"
            )
        except Exception as e:
            return fail(f"無法執行 twlegalrag 檢查指令：{e}")

        # 輸出原汁原味的 stdout 與 stderr
        sys.stdout.write(result.stdout)
        sys.stderr.write(result.stderr)

        # 解析文字以判定退出碼
        stdout_lower = result.stdout.lower()
        if "fail" in stdout_lower or "不在bundle/錯誤" in result.stdout or "高度疑似捏造" in result.stdout:
            return 1
        elif "needs_review" in stdout_lower or "待人工" in result.stdout:
            return 2
        else:
            return 0


if __name__ == "__main__":
    raise SystemExit(main())