"""
check_citations.py — 書狀草稿引用規則檢查員 (v1)

在 build_pleading.py 正式排版前，自動掃描 Markdown 草稿，
偵測六類引用規則違規，產出結構化報告。

規則代號：
  BARE_EVIDENCE     — 裸露的證據編號（前無名稱說明，後無「為憑」「可證」等收尾）
  EVIDENCE_NO_QUOTE — 援引書證內容但缺乏原文摘錄（無引號「…」包裹）
  TRANSCRIPT_NO_QUOTE — 引用筆錄/證詞但缺乏原文或問答紀錄
  RULING_NO_QUOTE   — 引用裁判字號但未附原文摘錄
  RULING_PARAPHRASE — 裁判引用疑似改寫（引號內含概述標誌詞）
  MISSING_PAGE_REF  — 引用卷宗/筆錄/證據但無頁碼也無[待補]佔位符

返回：
  violations (list[dict]) — 每則違規的完整資訊
  warnings   (list[dict]) — 需要人工核對的提示（不阻斷排版）
"""

import re
import sys
import os
from pathlib import Path

# 強制 stdout 使用 UTF-8（解決 Windows cp950 編碼問題）
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


# ─────────────────────────────────────────────
# 正規式設定
# ─────────────────────────────────────────────

# 偵測各類證據編號：甲證1、乙證2、上證3、被上證4、原證5、被證6…等
_EVIDENCE_TAG = re.compile(
    r'（(?:甲|乙|丙|丁|上|被上|原|被|附)證(?:\d+|[一二三四五六七八九十百]+)）'
)

# 「為憑」「可證」「可參」「足證」等引用收尾詞
_EVIDENCE_CLOSING = re.compile(r'(?:為憑|可證|可參|足證|參照|足認|足資認定|得認|有據|在卷可查)[。，]?')

# 援引書證內容的觸發詞（表示 AI 正在引用書證的實際內容）
_EVIDENCE_CONTENT_TRIGGER = re.compile(
    r'(?:依據|觀諸|參諸|依照|查閱|見|詳見|如|依|按)'
    r'(?:甲|乙|丙|丁|上|被上|原|被|附)證\d+'
    r'(?:記載|所載|所示|顯示|可知|所證明|足證|清楚顯示|明確載明|記錄)'
)

# 筆錄/證詞引用觸發詞
_TRANSCRIPT_TRIGGER = re.compile(
    r'(?:證人|被告|原告|上訴人|被上訴人|相對人|聲請人)'
    r'(?:[^\n，。]{0,10}?)'
    r'(?:證述|供稱|陳稱|陳述|作證|到庭證稱|證詞|作證表示|證稱)'
)

# 裁判字號識別（最高法院、高院、地院、行政法院等）
_RULING_CITE = re.compile(
    r'(?:最高法院|臺灣高等法院|高等法院|臺中高分院|高院|地方法院|臺灣.*?法院|行政法院|最高行政法院)'
    r'\s*'
    r'\d{2,3}\s*年度?\s*(?:台|臺)?[\w]+字第\s*\d+\s*號'
)

# 引號包裹的原文（「…」或『…』，允許換行）
_QUOTED_TEXT = re.compile(r'[「『][\s\S]{1,1000}?[」』]')

# 疑似概述/改寫的標誌詞
_PARAPHRASE_MARKERS = re.compile(
    r'(?:大意|意旨略以|旨在|意旨略謂|略以|概謂|旨趣|大略為|要旨略以|大意略以)'
)

# 卷宗、筆錄或書狀引用（必須附頁碼）
_VOLUME_REF = re.compile(
    r'(?:鈞院卷|一審卷|二審卷|原審卷|更審卷|前審卷|本院卷|北簡字卷|筆錄|'
    r'(?:原告|被告|上訴人|被上訴人|相對人|聲請人)(?:之)?(?:起訴狀|答辯狀|上訴狀|理由狀|準備書狀|陳報狀))'
)

# 頁碼偵測（第N頁、頁次N、卷第N頁、第N、M、P頁等多頁格式、見...第N頁）
_PAGE_NUMBER = re.compile(
    r'第\s*\d+(?:[、，,]\s*\d+)*\s*頁'   # 第92、97頁 / 第92頁
    r'|頁次\s*\d+'                         # 頁次23
    r'|\[待補\]'                            # 佔位符
    r'|卷第\s*\d+'                          # 卷第23頁
    r'|見.*?第\s*\d+\s*頁'                  # 見一審卷第XX頁
)

# 問答格式（問：/證人：/答：/A：/Q：）
_QA_FORMAT = re.compile(r'^(?:問|答|證人|A|Q)[：:]', re.MULTILINE)


# ─────────────────────────────────────────────
# 核心掃描邏輯
# ─────────────────────────────────────────────

def _surrounding_text(lines: list[str], lineno: int, window: int = 3) -> str:
    """取得指定行前後 window 行的合併文字，用於上下文判斷。"""
    start = max(0, lineno - window)
    end = min(len(lines), lineno + window + 1)
    return '\n'.join(lines[start:end])


def check_draft(md_path: str) -> tuple[list[dict], list[dict]]:
    """
    掃描指定的 Markdown 草稿，回傳：
      violations: list[dict]  — ❌ 必須修正的違規
      warnings:   list[dict]  — ⚠️ 建議人工核對的提示
    """
    path = Path(md_path)
    if not path.exists():
        raise FileNotFoundError(f"草稿檔案不存在: {md_path}")

    text = path.read_text(encoding='utf-8')
    lines = text.splitlines()

    violations: list[dict] = []
    warnings: list[dict] = []

    def add_violation(rule: str, lineno: int, excerpt: str, suggestion: str):
        violations.append({
            'rule': rule,
            'lineno': lineno + 1,   # 轉為 1-indexed
            'excerpt': excerpt.strip()[:120],
            'suggestion': suggestion,
        })

    def add_warning(rule: str, lineno: int, excerpt: str, note: str):
        warnings.append({
            'rule': rule,
            'lineno': lineno + 1,
            'excerpt': excerpt.strip()[:120],
            'note': note,
        })

    for i, line in enumerate(lines):
        stripped = line.strip()
        ctx = _surrounding_text(lines, i, window=4)

        # ── 規則 1: BARE_EVIDENCE ──────────────────────────────────────────
        # 偵測裸露的證據編號：前後文缺少名稱說明或正式收尾詞
        for m in _EVIDENCE_TAG.finditer(line):
            tag = m.group(0)
            before = line[:m.start()]
            after = line[m.end():]

            # 判斷前文是否有說明詞（至少 3 字的中文說明）
            has_desc_before = len(before.strip()) >= 3 and not before.strip().endswith('（')

            # 判斷後文是否有收尾詞
            has_closing = bool(_EVIDENCE_CLOSING.search(after[:30]))

            # 判斷是否為「純清單行」（如：甲證1：〇〇文件）
            is_list_line = bool(re.match(r'^[（甲乙丙丁上被原附](附件)?\d*[：:]', stripped))

            if not is_list_line and not (has_desc_before and has_closing):
                if not has_desc_before:
                    add_violation(
                        'BARE_EVIDENCE', i,
                        stripped,
                        f'在 {tag} 之前應先說明證據名稱，例如：「……此有○○文件{tag}為憑。」'
                    )
                elif not has_closing:
                    add_violation(
                        'BARE_EVIDENCE', i,
                        stripped,
                        f'{tag} 結尾應加上「為憑」、「可證」或「可參」等收尾詞。'
                    )

        # ── 規則 2: EVIDENCE_NO_QUOTE ──────────────────────────────────────
        # 援引書證內容卻沒有引號包裹
        if _EVIDENCE_CONTENT_TRIGGER.search(line):
            if not _QUOTED_TEXT.search(ctx):
                add_violation(
                    'EVIDENCE_NO_QUOTE', i,
                    stripped,
                    '引用書證內容時，應以「……」引號摘錄書證原文（如記載文字、合約條款等）。'
                )

        # ── 規則 3: TRANSCRIPT_NO_QUOTE ────────────────────────────────────
        # 引用筆錄/證詞但無引號或問答格式
        if _TRANSCRIPT_TRIGGER.search(line):
            has_quote_in_ctx = bool(_QUOTED_TEXT.search(ctx))
            has_qa_in_ctx = bool(_QA_FORMAT.search(ctx))
            if not has_quote_in_ctx and not has_qa_in_ctx:
                add_violation(
                    'TRANSCRIPT_NO_QUOTE', i,
                    stripped,
                    '引用證人證詞或筆錄時，應以「……」引號摘錄原文，'
                    '或以「問：…」/「證人：…」格式呈現逐字問答。'
                )

        # ── 規則 4: RULING_NO_QUOTE ────────────────────────────────────────
        # 裁判字號後缺乏引號原文
        if _RULING_CITE.search(line):
            if not _QUOTED_TEXT.search(ctx):
                add_violation(
                    'RULING_NO_QUOTE', i,
                    stripped,
                    '引用裁判時，必須在字號後方以「……」引號摘錄裁判書原文。'
                    '若來源僅有大意，請標註：（來源僅載大意，原文待核）。'
                )

        # ── 規則 5: RULING_PARAPHRASE ──────────────────────────────────────
        # 裁判引用有引號但疑似改寫（含概述標誌詞）
        if _RULING_CITE.search(line):
            for qm in _QUOTED_TEXT.finditer(ctx):
                if _PARAPHRASE_MARKERS.search(qm.group(0)):
                    add_warning(
                        'RULING_PARAPHRASE', i,
                        stripped,
                        '此裁判引文中含有「大意」「意旨略以」等概述詞，'
                        '可能並非裁判書逐字原文，請人工核對後確認。'
                    )
                    break

        # ── 規則 6: MISSING_PAGE_REF ───────────────────────────────────────
        # 引用卷宗、筆錄或書狀必須有頁碼或佔位符
        if _VOLUME_REF.search(line):
            # 排除純清單行
            is_list_line = bool(re.match(r'^[（甲乙丙丁上被原附](附件)?\d*[：:]', stripped))
            if not is_list_line and not _PAGE_NUMBER.search(ctx):
                add_violation(
                    'MISSING_PAGE_REF', i,
                    stripped,
                    '引用卷宗、筆錄或書狀內容時，請標示頁碼（如：卷第23頁），'
                    '或在頁碼未知時加入佔位符：第 [待補] 頁。'
                )

    return violations, warnings


# ─────────────────────────────────────────────
# 報告格式化
# ─────────────────────────────────────────────

def format_report(violations: list[dict], warnings: list[dict], md_path: str) -> str:
    """格式化為人類可讀的報告文字。"""
    lines_out = []
    total = len(violations)
    total_w = len(warnings)
    filename = os.path.basename(md_path)

    if total == 0 and total_w == 0:
        return f"[PASS] 草稿引用規則檢查通過（{filename}），無違規。"

    lines_out.append(f"[WARN] 草稿引用規則檢查報告（{filename}）")
    lines_out.append(f"       共發現 {total} 處必須修正違規、{total_w} 處建議核對提示")
    lines_out.append("-" * 55)

    for v in violations:
        lines_out.append(f"\n[FAIL] [{v['rule']}]  第 {v['lineno']} 行")
        lines_out.append(f"  原文：「{v['excerpt']}」")
        lines_out.append(f"  建議：{v['suggestion']}")

    if warnings:
        lines_out.append("\n" + "-" * 55)
        for w in warnings:
            lines_out.append(f"\n[INFO] [{w['rule']}]  第 {w['lineno']} 行")
            lines_out.append(f"  原文：「{w['excerpt']}」")
            lines_out.append(f"  提示：{w['note']}")

    lines_out.append("\n" + "-" * 55)
    if total > 0:
        lines_out.append("請修正上述 [FAIL] 違規後，再次要求產出 Word 檔。")
    else:
        lines_out.append("[INFO] 提示項目無需強制修正，但建議人工確認後再產出 Word 檔。")

    return '\n'.join(lines_out)


# ─────────────────────────────────────────────
# 主程式（供 build_pleading.py 呼叫）
# ─────────────────────────────────────────────

def run_check(md_path: str, strict: bool = True) -> bool:
    """
    執行引用規則檢查。
    - strict=True  → 有 violations 即回傳 False（阻斷排版）
    - strict=False → 僅印出報告，不阻斷
    回傳 True 表示可繼續排版，False 表示應中止。
    """
    violations, warnings = check_draft(md_path)
    report = format_report(violations, warnings, md_path)
    print(report)

    if strict and violations:
        return False
    return True


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python check_citations.py <draft.md>")
        sys.exit(1)

    md_file = sys.argv[1]
    ok = run_check(md_file, strict=True)
    sys.exit(0 if ok else 1)
