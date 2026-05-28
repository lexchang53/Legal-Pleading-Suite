from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


def fail(message: str, code: int = 1) -> int:
    print(message, file=sys.stderr)
    return code


def main() -> int:
    if len(sys.argv) != 3:
        return fail("用法：python scripts/check.py <bundle.json> <answer.txt>")

    bundle_path = Path(sys.argv[1]).expanduser().resolve()
    answer_path = Path(sys.argv[2]).expanduser().resolve()

    if shutil.which("twlegalrag") is None:
        return fail("找不到 twlegalrag 指令。請先安裝：pip install twlegalrag")

    if not bundle_path.exists():
        return fail(f"找不到 bundle 檔案：{bundle_path}")

    if not answer_path.exists():
        return fail(f"找不到回答檔案：{answer_path}")

    result = subprocess.run(
        ["twlegalrag", "check", str(bundle_path), str(answer_path)],
        text=True,
    )
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())