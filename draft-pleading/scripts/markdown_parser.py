# markdown_parser.py

"""
專門處理純文字與 Markdown 的解析與自癒（書狀領域專用 AST Parser）。
API 邊界：本模組對外僅暴露高階介面與資料契約，將 AST 建構細節封裝。
"""

import sys
import os
import re

# ==============================================================================
# 資料結構 (穩定公開的 DTO 契約)
# ==============================================================================

class TableBlock:
    def __init__(self, headers, rows):
        self.headers = headers
        self.rows = rows
        self.style = "TABLE"
        self.text = ""
        self.is_semantic_heading = False
        self.is_override_trigger = False
        self.needs_num = False
        self.ilvl = None
        self.has_l2_children = False
        self.is_pleading_heading = False

class Block:
    """表示 Markdown 中的一個段落區塊。"""
    def __init__(self, style, text, ilvl=None, needs_num=False,
                 is_override_trigger=False, raw_text=None,
                 is_pleading_heading=False, is_explicit_bold=False):
        self.style = style              # Word 樣式名稱
        self.text = text               # 去除前綴後的文字
        self.ilvl = ilvl               # 編號層級（0~3），若無編號則為 None
        self.needs_num = needs_num     # 是否需要設定 numPr
        self.is_override_trigger = is_override_trigger  # 是否觸發重新編號
        self.raw_text = raw_text or text  # 原始文字（含前綴）
        # 書狀相容層：由 #### 一、... 解析出的第一層主段標題
        # 此旗標控制：這個 通用_層級1 必須套用 heading-style bold
        self.is_pleading_heading = is_pleading_heading
        # 書狀相容層：此項是否為語意標題（聲明、理由等）
        self.is_semantic_heading = False
        # 書狀相容層：此 通用_層級1 是否有直屬的 通用_層級2 子段落
        # 由 post-processing pass 設定；預設 True（保守：不確定時仍粗體）
        self.has_l2_children = True
        # 是否含有 Markdown ** 顯式粗體標記
        self.is_explicit_bold = is_explicit_bold


# ==============================================================================
# 常數定義 (Parser 專用)
# ==============================================================================

# Markdown 層級前綴的正規表達式（優先順序由高到低）
# 支援並忽略包覆在最外層的 Markdown 粗體/斜體標記（** 或 __）
LEVEL_PATTERNS = [
    # (正則, 樣式名稱, ilvl)
    (re.compile(r'^(?:\*\*|__)?([一二三四五六七八九十百千]+)、\s*(.*?)(?:\*\*|__)?$'), '通用_層級1', 0),
    (re.compile(r'^(?:\*\*|__)?\(([一二三四五六七八九十百千]+)\)\s*(.*?)(?:\*\*|__)?$'), '通用_層級2', 1),
    (re.compile(r'^(?:\*\*|__)?(\d+)\.\s+(.*?)(?:\*\*|__)?$'), '通用_層級3', 2),
    (re.compile(r'^(?:\*\*|__)?\((\d+)\)\s*(.*?)(?:\*\*|__)?$'), '通用_層級4', 3),
]

# ---- 書狀相容層：H2 主標題（## 前綴）→ 書狀_標題
H2_PATTERN = re.compile(r'^#{2}\s+(.+)')

# ---- 書狀相容層：#### 主段標題（#### 一、...）→ 通用_層級1（書狀主段粗體用途）
#   匹配 #### 後接 CJK 數字+、 的第一層主段標題格式
H4_LEVEL1_PATTERN = re.compile(
    r'^#{4}\s+(?:\*\*|__)?([一二三四五六七八九十百千]+)、\s*(.*?)(?:\*\*|__)?$')

COMMENT_PATTERN = re.compile(r'^<!--\s*(.+?)\s*-->$')
# 帶值的指令（如 <!--法院: 文字-->）
COMMENT_KV_PATTERN = re.compile(r'^(.+?):\s*(.+)$')

# 動態證據編號
DYNAMIC_EVIDENCE_PATTERN = re.compile(r'^([^\d\s：:]+)(\d+)[：:](.*)$')

# 問答（blockquote）
QA_PATTERN = re.compile(r'^>\s*(問[：:].*|答[：:].*)$')


# ==============================================================================
# 內部輔助函式
# ==============================================================================

def _mark_children(blocks):
    """
    為所有層級（ilvl=0~3）的編號區塊設定 has_child。
    規則：向後掃描，如果在遇到相同或更上層的編號（ilvl <= current.ilvl）
    或結構性結束點之前，遇到了下屬層級（ilvl == current.ilvl + 1），即視為有子節點。
    """
    for i, block in enumerate(blocks):
        if not block.needs_num or block.ilvl is None:
            continue
            
        has_child = False
        for j in range(i + 1, len(blocks)):
            nxt = blocks[j]
            # 遇到編號段落
            if nxt.needs_num and nxt.ilvl is not None:
                if nxt.ilvl == block.ilvl + 1:
                    has_child = True
                    break
                elif nxt.ilvl <= block.ilvl:
                    # 遇到同級或上級，表示本段的子節點搜尋結束且無獲
                    break
                    
            # 遇到結構性中斷點（簽章、聲明等），結束搜尋
            if nxt.style in ('書狀_謹狀', '書狀_預設', '書狀_簽章', '書狀_標題', '書狀_狀首日期') or getattr(nxt, 'is_semantic_heading', False):
                # 但如果是單純的文字段落（書狀_預設），它可能是說明內文，並不是中斷點
                if nxt.style == '書狀_預設' and not getattr(nxt, 'is_semantic_heading', False):
                    continue
                break
                
        block.has_child = has_child


# ==============================================================================
# 公開 API
# ==============================================================================

def auto_fix_markdown(md_path: str) -> str:
    """
    讀取 Markdown 草稿，自動修正可機械修正的格式錯誤，並印出修正摘要。
    回傳修正後的文字（多行字串），同時在 stdout 印出修正摘要。
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    fixed_lines = []
    corrections = []  # [(行號, 說明, 原文, 修正後)]

    # ── 需要全文掃描的狀態追蹤 ──────────────────────────────────────────
    has_gongjian = any('公鑒' in ln for ln in lines)
    has_jinjing = any(ln.strip() in ('謹狀', '謹 狀') for ln in lines)

    fatal_errors = []
    if not has_gongjian:
        fatal_errors.append("草稿中缺少「公鑒」行，請確認書狀末尾是否有「[法院名稱]　公鑒」")
    if not has_jinjing:
        fatal_errors.append("草稿中缺少「謹狀」行，請確認書狀末尾是否有「謹狀」")

    if fatal_errors:
        print("\n[ERROR] ⛔ 偵測到無法自動修正的嚴重格式錯誤，排版終止：")
        for err in fatal_errors:
            print(f"  - {err}")
        print("\n請修正上述錯誤後重新提交。")
        sys.exit(2)

    seen_gongjian = False

    for lineno, raw in enumerate(lines, start=1):
        line = raw.rstrip('\n').rstrip('\r')
        original = line

        # ── 規則 1：誤用 ## / ### 作為聲明或理由區段標題 ──────────────
        _SEMANTIC_HEADINGS = {
            '訴之聲明', '聲明', '答辯聲明', '上訴聲明', '減縮後訴之聲明',
            '事實與理由', '理由', '答辯理由', '上訴理由',
        }
        m_hx = re.match(r'^(#{2,6})\s+(.+)$', line)
        if m_hx:
            heading_text = m_hx.group(2).strip().replace('**', '').replace('__', '')
            heading_no_space = heading_text.replace('\u3000', '').replace(' ', '')
            if heading_no_space in _SEMANTIC_HEADINGS:
                fixed = f'\u3000\u3000{heading_text}'
                corrections.append((lineno, '誤用 Markdown 標題改為全形空白前綴（語意標題）', original, fixed))
                line = fixed

        # ── 規則 2：理由段落（非狀首區）出現「民國」年份前綴 ──────────
        if not seen_gongjian:
            if re.search(r'民國\s*\d+\s*年', line):
                is_pure_date = re.match(r'^(民國)?\s*\d+\s*年\s*\d+\s*月\s*\d+\s*日\s*$', line.strip())
                is_interest_clause = bool(re.search(r'自.{0,30}民國\d+年', line))
                if not is_pure_date and not is_interest_clause:
                    fixed = re.sub(r'民國(\s*)(\d+\s*年)', r'\1\2', line)
                    if fixed != line:
                        corrections.append((lineno, '理由段落誤加「民國」紀元前綴，已移除', original, fixed))
                        line = fixed

        # ── 規則 3：具狀人行缺少半形冒號 ──────────────────────────────
        _SIGNATURE_ROLES = ['具狀人', '特別代理人', '訴訟代理人', '複代理人', '法定代理人', '選任辯護人', '代理人']
        for role in _SIGNATURE_ROLES:
            pat = re.compile(r'^' + re.escape(role) + r'[　 ]+([^：:].+)$')
            mm = pat.match(line.strip())
            if mm:
                fixed = f'{role}：{mm.group(1)}'
                corrections.append((lineno, f'「{role}」後方缺少冒號（：），已補上', original, fixed))
                line = fixed
                break

        # ── 規則 4：狀尾日期缺少「中華民國」前綴 ────────────────────
        if seen_gongjian:
            date_m = re.match(r'^(\d+)\s*年\s*\d+\s*月\s*\d+\s*日\s*$', line.strip())
            if date_m:
                fixed = f'中華民國{line.strip()}'
                corrections.append((lineno, '狀尾日期缺少「中華民國」完整前綴，已補上', original, fixed))
                line = fixed

        # ── 規則 5：「日期：NNN年M月D日」舊格式 → 脫去前綴並補「中華民國」────
        if seen_gongjian:
            date_prefix_m = re.match(
                r'^日期[：:]\s*(\d+\s*年\s*\d+\s*月\s*\d+\s*日)\s*$', line.strip()
            )
            if date_prefix_m:
                fixed = f'中華民國{date_prefix_m.group(1)}'
                corrections.append((lineno, '狀尾「日期：NNN年...」舊格式，已轉換為「中華民國NNN年...」', original, fixed))
                line = fixed

        if '公鑒' in line:
            seen_gongjian = True

        fixed_lines.append(line + '\n')

    if corrections:
        print(f"\n[AUTO-FIX] ⚠️  共自動修正 {len(corrections)} 處格式問題（排版繼續執行）：")
        for lineno, desc, orig, fixed in corrections:
            print(f"  行 {lineno:4d}：{desc}")
            print(f"           原文：{orig[:60]}{'...' if len(orig) > 60 else ''}")
            print(f"           修正：{fixed[:60]}{'...' if len(fixed) > 60 else ''}")
        print()
    else:
        print("[AUTO-FIX] ✅ 未偵測到可自動修正的格式問題\n")

    return ''.join(fixed_lines)


def parse_markdown(md_path, content=None):
    """
    解析 Markdown 檔案為 Block 清單。
    """
    if content is not None:
        lines = content.splitlines(keepends=True)
    else:
        with open(md_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

    blocks = []
    i = 0
    has_seen_title = False
    has_seen_gongjian = False
    seen_juzhuangren = False 
    in_table = False
    table_headers = []
    table_rows = []
    
    while i < len(lines):
        line = lines[i].rstrip('\n').rstrip('\r')
        i += 1

        if not line.strip():
            if in_table:
                blocks.append(TableBlock(table_headers, table_rows))
                in_table = False
                table_headers = []
                table_rows = []
            continue

        line_stripped = line.strip()
        no_space = line_stripped.replace('\u3000', '').replace(' ', '').replace('**', '').replace('__', '')

        if line_stripped.startswith('|') and line_stripped.endswith('|'):
            cells = [c.strip() for c in line_stripped.strip('|').split('|')]
            if not in_table:
                table_headers = cells
                in_table = True
            elif all(re.match(r'^[-:]+$', c.replace(' ', '')) for c in cells):
                continue
            else:
                table_rows.append(cells)
            continue
        else:
            if in_table:
                blocks.append(TableBlock(table_headers, table_rows))
                in_table = False
                table_headers = []
                table_rows = []

        _SEMANTIC_DECL_WORDS = {'訴之聲明', '聲明', '減縮後訴之聲明', '聲明事項', '反訴聲明', '上訴聲明', '答辯聲明', '抗告聲明', '追加聲明', '變更聲明'}
        _SEMANTIC_REASON_WORDS = {'事實與理由', '理由'}

        h2_m = H2_PATTERN.match(line.strip())
        if h2_m:
            h2_inner = h2_m.group(1).strip()
            h2_inner_no_space = h2_inner.replace('\u3000', '').replace(' ', '')
            if h2_inner_no_space in _SEMANTIC_DECL_WORDS:
                b = Block(style='書狀_預設', text=f'**　　{h2_inner}**')
                b.is_semantic_heading = True
                blocks.append(b)
                continue
            elif h2_inner_no_space in _SEMANTIC_REASON_WORDS:
                b = Block(style='書狀_預設', text=f'**　　{h2_inner}**', is_override_trigger=True)
                b.is_semantic_heading = True
                blocks.append(b)
                continue
            else:
                blocks.append(Block(style='書狀_標題', text=h2_inner))
                has_seen_title = True
                continue

        h4_m = H4_LEVEL1_PATTERN.match(line.strip())
        if h4_m:
            text_after = h4_m.group(2).strip()
            blocks.append(Block(
                style='通用_層級1', text=text_after,
                ilvl=0, needs_num=True,
                raw_text=line.strip(),
                is_pleading_heading=True   
            ))
            continue

        if line.startswith('　　') or line.startswith('**　　') or line.startswith('__　　'):
            if no_space.endswith('聲明') or no_space.endswith('理由') or no_space.endswith('事項') or no_space.endswith('聲明事項') or no_space.endswith('聲請事項'):
                clean_text = line.rstrip()
                if not clean_text.startswith('**') and not clean_text.startswith('__'):
                    clean_text = f'**{clean_text}**'
                b = Block(style='書狀_預設', text=clean_text, is_override_trigger=True)
                b.is_semantic_heading = True
                blocks.append(b)
                continue

        if line.startswith('# ') and not line.startswith('## '):
            text = line[2:].strip()
            blocks.append(Block(style='書狀_標題', text=text))
            has_seen_title = True
            continue
        elif not has_seen_title and not line.startswith('<!--') and not no_space.endswith('聲明') and no_space not in _SEMANTIC_REASON_WORDS:
            is_level = any(pat.match(line.strip()) for pat, _, _ in LEVEL_PATTERNS)
            if not is_level and ':' not in line and '：' not in line and '<' not in line:
                blocks.append(Block(style='書狀_標題', text=line.strip()))
                has_seen_title = True
                continue

        m = COMMENT_PATTERN.match(line.strip())
        if m:
            content = m.group(1)
            kv = COMMENT_KV_PATTERN.match(content)
            if kv:
                key, value = kv.group(1).strip(), kv.group(2).strip()
                if key == '法院':
                    blocks.append(Block(style='書狀_預設', text=f'{value}　公鑒'))
                elif key == '簽章':
                    for part in value.split(';'):
                        part = part.strip()
                        if part:
                            blocks.append(Block(style='書狀_簽章', text=part))
                elif key == '日期':
                    blocks.append(Block(style='書狀_簽章', text=f'日期：{value}'))
                elif key == '狀首日期':
                    blocks.append(Block(style='書狀_狀首日期', text=value))
                else:
                    blocks.append(Block(style='書狀_預設', text=content))
            elif content == '謹狀':
                blocks.append(Block(style='書狀_謹狀', text='謹狀'))
            elif content in ('聲明', '理由', '訴之聲明', '事實與理由'):
                blocks.append(Block(style='書狀_預設', text=content))
            else:
                blocks.append(Block(style='書狀_預設', text=content))
            continue

        qa_m = QA_PATTERN.match(line.strip())
        if qa_m:
            blocks.append(Block(
                style='書狀_清單', text=qa_m.group(1),
                needs_num=False
            ))
            continue

        ev_m = DYNAMIC_EVIDENCE_PATTERN.match(line.strip())
        if ev_m:
            prefix, num_str, desc = ev_m.group(1), ev_m.group(2), ev_m.group(3)
            num = int(num_str)
            if len(prefix) <= 10 and ('證' in prefix or '附' in prefix):
                if len(prefix) <= 2:
                    style = '書狀_證據編號10' if num >= 10 else '書狀_證據編號'
                else:
                    style = '書狀_被上證據編號10' if num >= 10 else '書狀_被上證據編號'
                blocks.append(Block(style=style, text=line.strip()))
                continue

        stripped = line.strip()
        if stripped in ('謹狀', '謹 狀'):
            blocks.append(Block(style='書狀_謹狀', text='謹狀'))
            continue
        if '公鑒' in stripped:
            has_seen_gongjian = True
            blocks.append(Block(style='書狀_預設', text=stripped))
            continue
            
        if re.match(r'^(具狀人|特別代理人|訴訟代理人|複代理人|選任辯護人|法定代理人|代理人)[:：]', stripped):
            norm_text = stripped
            if re.search(r'中華民國', stripped):
                norm_text = stripped.replace('中華民國', '民國')
            if re.match(r'^具狀人[:：]', stripped):
                seen_juzhuangren = True
            blocks.append(Block(style='書狀_簽章', text=norm_text))
            continue
            
        if has_seen_gongjian and seen_juzhuangren:
            generic_sig = re.match(r'^[\u4e00-\u9fff]{1,8}[:：](.+)$', stripped)
            if generic_sig:
                blocks.append(Block(style='書狀_簽章', text=stripped))
                continue

        if re.match(r'^中華民國\s*\d+\s*年', stripped) or re.match(r'^民國\s*\d+\s*年', stripped):
            norm_text = stripped.replace('中華民國', '民國')
            if has_seen_gongjian:
                blocks.append(Block(style='書狀_狀尾日期', text=norm_text))
            else:
                blocks.append(Block(style='書狀_簽章', text=norm_text))
            continue

        if has_seen_title and re.match(r'^(中華民國)?\s*(民國)?\s*\d+\s*年\s*\d+\s*月\s*\d+\s*日$', stripped):
            norm_date = re.sub(r'^(中華民國|民國)\s*', '', stripped).strip()
            blocks.append(Block(style='書狀_狀首日期', text=norm_date))
            continue

        matched = False
        for pattern, style, ilvl in LEVEL_PATTERNS:
            lm = pattern.match(line.strip())
            if lm:
                text = lm.group(lm.lastindex)
                is_explicit_bold = bool(re.search(r'\*\*|__', line))
                
                blocks.append(Block(
                    style=style, text=text,
                    ilvl=ilvl, needs_num=True,
                    raw_text=line.strip(),
                    is_explicit_bold=is_explicit_bold
                ))
                matched = True
                break

        if matched:
            continue

        blocks.append(Block(style='書狀_預設', text=line.strip(), raw_text=line))

    if in_table:
        blocks.append(TableBlock(table_headers, table_rows))

    # 隱藏層處理：標記是否有子層級
    _mark_children(blocks)

    if md_path:
        filename = os.path.basename(md_path)
    else:
        filename = 'memory_content'
    print(f"[INFO] 從 '{filename}' 解析出 {len(blocks)} 個段落區塊")
    
    return blocks
