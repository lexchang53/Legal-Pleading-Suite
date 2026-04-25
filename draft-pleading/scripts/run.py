#!/usr/bin/env python3
"""
Runner for draft-pleading scripts.
Uses notebooklm skill's virtual environment since batch_query.py
depends on notebooklm's dependencies (patchright, etc.).

Usage:
    python run.py batch_query.py --input queries.json --output results.json
"""

import os
import sys
import subprocess
from pathlib import Path


def get_notebooklm_venv_python():
    """Get the notebooklm skill's virtual environment Python executable"""
    skill_dir = Path(__file__).parent.parent
    notebooklm_dir = skill_dir.parent / "notebooklm-skill"
    venv_dir = notebooklm_dir / ".venv"

    if not venv_dir.exists():
        print("❌ NotebookLM skill's virtual environment not found.")
        print("   Please run a NotebookLM query first to set up the environment:")
        print("   python scripts/run.py auth_manager.py status")
        sys.exit(1)

    if os.name == 'nt':  # Windows
        venv_python = venv_dir / "Scripts" / "python.exe"
    else:
        venv_python = venv_dir / "bin" / "python"

    if not venv_python.exists():
        print(f"❌ Python executable not found: {venv_python}")
        sys.exit(1)

    return venv_python


def main():
    if len(sys.argv) < 2:
        print("Usage: python run.py <script_name> [args...]")
        print("\nAvailable scripts:")
        print("  batch_query.py  - Batch query NotebookLM in a single browser session")
        sys.exit(1)

    script_name = sys.argv[1]
    script_args = sys.argv[2:]

    # Ensure .py extension
    if not script_name.endswith('.py'):
        script_name += '.py'

    # Get script path
    script_dir = Path(__file__).parent
    script_path = script_dir / script_name

    if not script_path.exists():
        print(f"❌ Script not found: {script_name}")
        sys.exit(1)

    # Use notebooklm's venv Python
    venv_python = get_notebooklm_venv_python()

    # Build and run command
    cmd = [str(venv_python), str(script_path)] + script_args

    try:
        result = subprocess.run(cmd)
        sys.exit(result.returncode)
    except KeyboardInterrupt:
        print("\n⚠️ Interrupted by user")
        sys.exit(130)
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
