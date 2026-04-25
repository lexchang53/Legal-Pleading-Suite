#!/usr/bin/env python3
"""
Draft-Pleading Batch Query Script

在單一瀏覽器 session 中連續查詢多個 NotebookLM 問題，
避免每個問題都重新開關瀏覽器的時間開銷。

用法：
    python batch_query.py --input queries.json --output results.json
    python batch_query.py --input queries.json --output results.json --show-browser

輸入 JSON 格式：
{
    "notebook_id": "私人勞動法知識庫",
    "questions": [
        {"id": "q1", "text": "問題一..."},
        {"id": "q2", "text": "問題二..."}
    ]
}

輸出 JSON 格式：
{
    "notebook_id": "私人勞動法知識庫",
    "notebook_url": "https://...",
    "total_questions": 2,
    "successful": 2,
    "failed": 0,
    "results": [
        {
            "id": "q1", "question": "...", "answer": "...", "status": "success",
            "source_tier": "internal",
            "is_primary_authority": true,
            "has_case_id": null,
            "has_statute_ref": null,
            "verification_status": "pending",
            "draft_eligible": false,
            "risk_flags": []
        }
    ]
}

注意：
- 本腳本僅處理「第一層——內部知識庫主查詢」。
- 「第二層——定向外部搜尋」目前僅完成流程規範與資料欄位保留，
  待實作外部搜尋介面後接入。
- 來源分級欄位（source_tier 等）已預留於輸出格式中，
  以便後續與驗證報告與草稿生成流程整合。
"""

import sys
import json
import time
import argparse
from pathlib import Path

# Paths
SCRIPT_DIR = Path(__file__).parent
SKILL_DIR = SCRIPT_DIR.parent
NOTEBOOKLM_SKILL_DIR = SKILL_DIR.parent / "notebooklm-skill"
NOTEBOOKLM_SCRIPTS_DIR = NOTEBOOKLM_SKILL_DIR / "scripts"
KB_PROMPT_FILE = SKILL_DIR / "references" / "kb-prompt.md"

# Add notebooklm scripts to path for imports
sys.path.insert(0, str(NOTEBOOKLM_SCRIPTS_DIR))

from patchright.sync_api import sync_playwright
from browser_utils import BrowserFactory, StealthUtils
from browser_session import BrowserSession
from notebook_manager import NotebookLibrary
from auth_manager import AuthManager


def load_kb_prompt() -> str:
    """Load the knowledge base prompt prefix from kb-prompt.md"""
    if not KB_PROMPT_FILE.exists():
        print(f"⚠️ kb-prompt.md not found at {KB_PROMPT_FILE}")
        return ""
    
    with open(KB_PROMPT_FILE, "r", encoding="utf-8") as f:
        return f.read().strip()


def load_queries(input_file: str) -> dict:
    """Load queries from JSON file"""
    with open(input_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    if "notebook_id" not in data:
        raise ValueError("Missing 'notebook_id' in input JSON")
    if "questions" not in data or not data["questions"]:
        raise ValueError("Missing or empty 'questions' in input JSON")
    
    return data


def resolve_notebook_url(notebook_id: str) -> str:
    """Resolve notebook ID/name to URL using the library"""
    library = NotebookLibrary()
    notebook = library.get_notebook(notebook_id)
    
    if not notebook:
        # Try searching by name
        results = library.search_notebooks(notebook_id)
        if results:
            notebook = results[0]
        else:
            raise ValueError(
                f"Notebook '{notebook_id}' not found. "
                f"Run: python scripts/run.py notebook_manager.py list"
            )
    
    return notebook["url"]


def run_batch_queries(
    input_file: str,
    output_file: str,
    headless: bool = True,
    prepend_kb_prompt: bool = True
):
    """
    Run batch queries against a NotebookLM notebook in a single browser session.
    
    Args:
        input_file: Path to input JSON with questions
        output_file: Path to output JSON for results  
        headless: Run browser headlessly
        prepend_kb_prompt: Whether to prepend kb-prompt.md content to each question
    """
    # Check auth
    auth = AuthManager()
    if not auth.is_authenticated():
        print("⚠️ Not authenticated. Run: python scripts/run.py auth_manager.py setup")
        sys.exit(1)
    
    # Load inputs
    data = load_queries(input_file)
    notebook_id = data["notebook_id"]
    questions = data["questions"]
    
    print(f"📋 Loaded {len(questions)} questions for notebook: {notebook_id}")
    
    # Resolve notebook URL
    notebook_url = resolve_notebook_url(notebook_id)
    print(f"📚 Notebook URL: {notebook_url}")
    
    # Load KB prompt
    kb_prompt = ""
    if prepend_kb_prompt:
        kb_prompt = load_kb_prompt()
        if kb_prompt:
            print(f"📝 Loaded kb-prompt.md ({len(kb_prompt)} chars)")
        else:
            print("⚠️ No kb-prompt.md found, questions will be sent as-is")
    
    # Results collection
    results = []
    playwright = None
    context = None
    session = None
    
    try:
        # Start one browser session for all queries
        print("\n🚀 Starting browser session...")
        start_time = time.time()
        
        playwright = sync_playwright().start()
        context = BrowserFactory.launch_persistent_context(
            playwright, headless=headless
        )
        
        session = BrowserSession(
            session_id="batch-query",
            context=context,
            notebook_url=notebook_url
        )
        
        browser_time = time.time() - start_time
        print(f"✅ Browser ready in {browser_time:.1f}s (this cost is paid only once!)\n")
        
        # Query each question sequentially in the same session
        for i, q in enumerate(questions, 1):
            q_id = q.get("id", f"q{i}")
            q_text = q["text"]
            
            # Prepend KB prompt if available
            full_question = f"{kb_prompt}\n\n{q_text}" if kb_prompt else q_text
            
            print(f"━━━ Question {i}/{len(questions)} [{q_id}] ━━━")
            print(f"  📝 {q_text[:80]}...")
            
            q_start = time.time()
            result = session.ask(full_question)
            q_elapsed = time.time() - q_start
            
            result["id"] = q_id
            result["original_question"] = q_text  # Without KB prompt prefix
            result["elapsed_seconds"] = round(q_elapsed, 1)
            
            # === 來源分級欄位（內部知識庫查詢結果預設值） ===
            # 注意：目前本腳本僅處理第一層（內部知識庫主查詢）。
            # 第二層（定向外部搜尋）的結果將由專屬模組處理，
            # 待實作搜尋介面後接入。屆時外部搜尋結果將分流儲存，
            # 不會與內部知識庫結果混寫、混排或混送入草稿。
            result["source_tier"] = "internal"  # 內部知識庫；外部結果將為 "tier1"/"tier2"/"tier3"
            result["is_primary_authority"] = True  # 內部知識庫視為主查詢權威來源
            result["has_case_id"] = None  # 待 AI 代理後續分析填入
            result["has_statute_ref"] = None  # 待 AI 代理後續分析填入
            result["verification_status"] = "pending"  # pending / verified / rejected
            result["draft_eligible"] = False  # 須經驗證報告與使用者確認後才設為 True
            result["risk_flags"] = []  # 待 AI 代理填入風險標記
            
            results.append(result)
            
            status_icon = "✅" if result["status"] == "success" else "❌"
            print(f"  {status_icon} [{q_elapsed:.1f}s] {result['status']}")
            
            # Brief pause between queries to avoid rate limiting
            if i < len(questions):
                StealthUtils.random_delay(2000, 4000)
        
        total_time = time.time() - start_time
        successful = sum(1 for r in results if r["status"] == "success")
        failed = len(results) - successful
        
        print(f"\n{'━' * 50}")
        print(f"📊 Batch complete: {successful} success, {failed} failed")
        print(f"⏱️ Total time: {total_time:.1f}s")
        print(f"   (Browser startup: {browser_time:.1f}s, "
              f"avg per question: {(total_time - browser_time) / len(questions):.1f}s)")
        
    except Exception as e:
        print(f"\n❌ Batch query error: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Clean up
        if session:
            try:
                session.close()
            except:
                pass
        if context:
            try:
                context.close()
            except:
                pass
        if playwright:
            try:
                playwright.stop()
            except:
                pass
    
    # Write results
    output_data = {
        "notebook_id": notebook_id,
        "notebook_url": notebook_url,
        "total_questions": len(questions),
        "successful": sum(1 for r in results if r["status"] == "success"),
        "failed": sum(1 for r in results if r["status"] != "success"),
        "results": results
    }
    
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    
    print(f"\n💾 Results saved to: {output_file}")
    return output_data


def main():
    parser = argparse.ArgumentParser(
        description="Batch query NotebookLM in a single browser session"
    )
    parser.add_argument(
        "--input", required=True,
        help="Path to input JSON file with questions"
    )
    parser.add_argument(
        "--output", required=True,
        help="Path to output JSON file for results"
    )
    parser.add_argument(
        "--show-browser", action="store_true",
        help="Show browser window (for debugging)"
    )
    parser.add_argument(
        "--no-kb-prompt", action="store_true",
        help="Do not prepend kb-prompt.md to questions"
    )
    
    args = parser.parse_args()
    
    run_batch_queries(
        input_file=args.input,
        output_file=args.output,
        headless=not args.show_browser,
        prepend_kb_prompt=not args.no_kb_prompt
    )


if __name__ == "__main__":
    main()
