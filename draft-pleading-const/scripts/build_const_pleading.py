#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import subprocess
import sys
from pathlib import Path


def run_step(command):
    subprocess.run(command, check=True)


def existing_path(path, label):
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")
    return path


def main():
    parser = argparse.ArgumentParser(
        description="Build a constitutional court pleading DOCX and ODT from Markdown."
    )
    parser.add_argument("draft", help="Input Markdown draft path.")
    parser.add_argument(
        "--docx",
        help="Output DOCX path. Defaults to the draft path with a .docx suffix.",
    )
    parser.add_argument(
        "--odt",
        help="Output ODT path. Defaults to the DOCX path with a .odt suffix.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=180,
        help="Timeout, in seconds, for the DOCX to ODT conversion step.",
    )
    args = parser.parse_args()

    draft = existing_path(Path(args.draft).expanduser().resolve(), "Draft")
    docx = Path(args.docx).expanduser().resolve() if args.docx else draft.with_suffix(".docx")
    odt = Path(args.odt).expanduser().resolve() if args.odt else docx.with_suffix(".odt")

    skill_dir = Path(__file__).resolve().parents[1]
    skills_root = skill_dir.parent

    draft_pleading = skills_root / "draft-pleading"
    docx_to_odt = skills_root / "docx-to-odt"

    build_pleading = existing_path(
        draft_pleading / "scripts" / "build_pleading.py",
        "draft-pleading build script",
    )
    template = existing_path(
        draft_pleading / "assets" / "pleading-tmpl.docx",
        "draft-pleading template",
    )
    postprocess = existing_path(
        skill_dir / "scripts" / "postprocess_const_pleading.py",
        "constitutional pleading postprocess script",
    )
    convert = existing_path(
        docx_to_odt / "scripts" / "convert_docx_to_odt.py",
        "docx-to-odt conversion script",
    )

    docx.parent.mkdir(parents=True, exist_ok=True)
    odt.parent.mkdir(parents=True, exist_ok=True)

    run_step([
        sys.executable,
        str(build_pleading),
        str(draft),
        "--template",
        str(template),
        "--output",
        str(docx),
    ])
    run_step([sys.executable, str(postprocess), str(docx)])
    run_step([
        sys.executable,
        str(convert),
        str(docx),
        "--output",
        str(odt),
        "--timeout",
        str(args.timeout),
    ])

    print(f"[OK] DOCX: {docx}")
    print(f"[OK] ODT: {odt}")


if __name__ == "__main__":
    main()
