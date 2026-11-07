#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
odt_to_md.py — 將書狀 ODT（或 DOCX）轉回 draft-pleading 格式的 Markdown 草稿

用途：使用者在 Writer／Word 直接修改書狀後，把修改後的檔案轉回 md，
讓修改補充模式可以繼續以 md 為準。輸出格式依 references/phase5-draft-rules.md。

用法：
    python odt_to_md.py <書狀.odt|書狀.docx> [--output 路徑] [--diff 舊草稿.md] [--force]

    預設輸出：與來源同資料夾、同檔名，副檔名改為 .md。
    既有同名 .md 不覆蓋，除非加 --force。來源檔一律唯讀，不會被修改。
    --diff：轉換後列出與指定 md 的差異（忽略空白行）。

編號重建：
    排版引擎產出的檔案中，「一、(一) 1. (1)」是 Writer／Word 自動編號，文字裡沒有。
    本工具依段落樣式「通用_層級1～4」決定層級，上層出現時下層歸零；
    另支援 Writer 的 <text:list> 巢狀清單、start-value、手動取消編號。
    遇「理由／事實與理由」等區塊標題或「壹、貳、」大標題時全部歸零（與 build_pleading.py 一致）。
"""

import argparse
import difflib
import os
import re
import sys
import zipfile

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

# ----------------------------------------------------------------------------
# 共用常數
# ----------------------------------------------------------------------------

LEVEL_STYLES = {'通用_層級1': 0, '通用_層級2': 1, '通用_層級3': 2, '通用_層級4': 3}

DECL_WORDS = {'訴之聲明', '聲明', '減縮後訴之聲明', '聲明事項', '反訴聲明', '上訴聲明', '答辯聲明',
              '抗告聲明', '追加聲明', '變更聲明', '聲請審查客體', '應受判決事項之聲明', '主要爭點'}
REASON_WORDS = {'事實與理由', '理由', '聲請理由'}

MAJOR_HEADING = re.compile(r'^[壹貳參肆伍陸柒捌玖拾]+、')
DATE_ONLY = re.compile(r'^(?:中華民國|民國)?\s*\d+\s*年\s*\d+\s*月\s*\d+\s*日$')
ROC_DATE = re.compile(r'^中華民國\s*\d+\s*年')
SIGNATURE = re.compile(r'^[一-鿿]{1,8}[：:]')
EVIDENCE_LINE = re.compile(r'^([^\d\s：:]+)(\d+)[：:]')
INTRO_LINE = re.compile(r'^(為就|為).*事[：:]\s*$')
REPRESENTATIVE = re.compile(r'(代理人|辯護人|輔佐人)$')

CN_DIGITS = '零一二三四五六七八九'


def to_cjk(n: int) -> str:
    """1～99 的國字數字（十、十一、二十、二十一……）。"""
    if n <= 0:
        return str(n)
    if n < 10:
        return CN_DIGITS[n]
    if n < 100:
        tens, ones = divmod(n, 10)
        s = ('' if tens == 1 else CN_DIGITS[tens]) + '十'
        return s + (CN_DIGITS[ones] if ones else '')
    return str(n)


def number_prefix(level: int, n: int) -> str:
    if level == 0:
        return f'{to_cjk(n)}、'
    if level == 1:
        return f'({to_cjk(n)}) '
    if level == 2:
        return f'{n}. '
    return f'({n}) '


def decode_odf_name(name):
    if not name:
        return name
    return re.sub(r'_([0-9a-fA-F]{2})_', lambda m: chr(int(m.group(1), 16)), name)


# ----------------------------------------------------------------------------
# 段落模型
# ----------------------------------------------------------------------------

class Para:
    """kind: 'title' | 'party' | 'level' | 'body' | 'block' | 'qa' | 'table' | 'empty'"""

    def __init__(self, text, style, level=None, numbered=True, start_value=None, list_restart=False):
        self.text = text
        self.style = style or ''
        self.level = level
        self.numbered = numbered
        self.start_value = start_value
        self.list_restart = list_restart


# ----------------------------------------------------------------------------
# ODT 讀取
# ----------------------------------------------------------------------------

def read_odt(path):
    from lxml import etree
    ns_text = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
    ns_style = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
    ns_office = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
    ns_table = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
    T = '{%s}' % ns_text
    S = '{%s}' % ns_style

    with zipfile.ZipFile(path) as z:
        content = etree.fromstring(z.read('content.xml'))
        styles = etree.fromstring(z.read('styles.xml'))

    # 樣式表：name -> (parent, list-style-name 是否被明確設為空)
    table = {}
    for root in (styles, content):
        for st in root.iter(S + 'style'):
            if st.get(S + 'family') != 'paragraph':
                continue
            name = st.get(S + 'name')
            lsn = st.get(S + 'list-style-name')
            table[name] = (st.get(S + 'parent-style-name'), lsn)

    def resolve(style_name):
        """回傳 (可辨識的樣式名稱, 是否取消編號)。沿 parent 鏈往上找到已知樣式。"""
        no_num = False
        cur = style_name
        seen = set()
        first = True
        while cur and cur not in seen:
            seen.add(cur)
            decoded = decode_odf_name(cur)
            parent, lsn = table.get(cur, (None, None))
            if first and lsn == '':
                no_num = True
            first = False
            if decoded in LEVEL_STYLES or decoded.startswith('書狀_') or decoded in ('Standard',):
                return decoded, no_num
            cur = parent
        return decode_odf_name(style_name), no_num

    def text_of(el):
        out = []

        def rec(e):
            if e.text:
                out.append(e.text)
            for c in e:
                tag = etree.QName(c).localname
                if tag == 'tab':
                    out.append('\t')
                elif tag == 's':
                    out.append(' ' * int(c.get(T + 'c', '1') or 1))
                elif tag == 'line-break':
                    out.append('\n')
                elif tag in ('note', 'annotation', 'tracked-changes', 'deletion', 'bookmark',
                             'bookmark-start', 'bookmark-end', 'soft-page-break', 'sequence-decls'):
                    pass
                else:
                    rec(c)
                if c.tail:
                    out.append(c.tail)

        rec(el)
        return ''.join(out)

    paras = []
    body = content.find('.//{%s}text' % ns_office)

    def walk(el, list_depth, in_header_item, list_style_stack):
        for c in el:
            tag = etree.QName(c).localname
            ns = etree.QName(c).namespace
            if ns == ns_text and tag == 'list':
                restart = (list_depth == 0 and c.get(T + 'continue-numbering') != 'true'
                           and not c.get(T + 'continue-list'))
                style = c.get(T + 'style-name') or (list_style_stack[-1] if list_style_stack else None)
                if restart and paras is not None:
                    paras.append(Para('', '__list_start__', list_restart=True))
                walk(c, list_depth + 1, False, list_style_stack + [style])
            elif ns == ns_text and tag in ('list-item', 'list-header'):
                sv = c.get(T + 'start-value')
                if sv:
                    paras.append(Para('', '__start_value__', level=list_depth - 1, start_value=int(sv)))
                walk(c, list_depth, tag == 'list-header', list_style_stack)
            elif ns == ns_text and tag in ('p', 'h'):
                sname, no_num = resolve(c.get(T + 'style-name'))
                txt = text_of(c)
                level = LEVEL_STYLES.get(sname)
                if level is None and list_depth > 0 and not in_header_item:
                    level = list_depth - 1
                numbered = level is not None and not no_num and not in_header_item
                paras.append(Para(txt, sname, level=level, numbered=numbered))
            elif ns == ns_table and tag == 'table':
                paras.append(Para('[表格：odt_to_md 未轉換，請人工處理]', '__table__'))
            elif tag == 'section':
                walk(c, list_depth, in_header_item, list_style_stack)

    walk(body, 0, False, [])
    return paras


# ----------------------------------------------------------------------------
# DOCX 讀取
# ----------------------------------------------------------------------------

def read_docx(path):
    import docx
    from docx.oxml.ns import qn
    d = docx.Document(path)
    paras = []
    for child in d.element.body.iterchildren():
        tag = child.tag.split('}')[-1]
        if tag == 'tbl':
            paras.append(Para('[表格：odt_to_md 未轉換，請人工處理]', '__table__'))
            continue
        if tag != 'p':
            continue
        p = docx.text.paragraph.Paragraph(child, d)
        sname = p.style.name if p.style is not None else ''
        # 沿 base_style 找已知樣式
        st = p.style
        while st is not None and st.name not in LEVEL_STYLES and not st.name.startswith('書狀_'):
            st = st.base_style
        if st is not None:
            sname = st.name
        out = []
        for node in child.iter():
            t = node.tag.split('}')[-1]
            if t == 't' and node.text:
                out.append(node.text)
            elif t == 'tab' and node.getparent().tag.endswith('}r'):
                out.append('\t')
            elif t in ('br', 'cr'):
                out.append('\n')
        txt = ''.join(out)
        level = LEVEL_STYLES.get(sname)
        no_num = False
        pPr = child.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                numId = numPr.find(qn('w:numId'))
                if numId is not None and numId.get(qn('w:val')) == '0':
                    no_num = True
        paras.append(Para(txt, sname, level=level, numbered=level is not None and not no_num))
    return paras


# ----------------------------------------------------------------------------
# 轉換
# ----------------------------------------------------------------------------

def norm_inline(s):
    """一般段落：去除換行與多餘空白（保留英文單字間單一空白與全形空白）。"""
    s = s.replace('\n', '')
    s = s.replace('\t', ' ')
    s = re.sub(r' {2,}', ' ', s)
    return s.strip(' ')


def norm_party(s):
    """狀首當事人：Tab 與換行合併為單一空白；頓號後的換行直接接續。"""
    s = re.sub(r'\n[\t ]*', '\n', s)
    s = re.sub(r'([、，,])\n', r'\1', s)
    s = s.replace('\n', ' ').replace('\t', ' ')
    s = re.sub(r' {2,}', ' ', s)
    return s.strip()


def party_role(s):
    m = re.match(r'^([^\s\t]+)', s)
    return m.group(1) if m else ''


def is_block_heading(text):
    core = text.replace('*', '').replace('　', '').replace(' ', '').strip()
    return core in DECL_WORDS or core in REASON_WORDS


def convert(paras):
    lines = []
    warnings = []

    # ---- 切分：狀首 / 正文 / 狀尾 ----
    idx_intro = None
    for i, p in enumerate(paras):
        if p.style.startswith('__'):
            continue
        if INTRO_LINE.match(norm_inline(p.text)) and p.level is None:
            idx_intro = i
            break
    if idx_intro is None:
        warnings.append('找不到「為就……事：」一行，狀首與正文的分界改以第一個編號段落判斷。')
        idx_intro = next((i for i, p in enumerate(paras) if p.level is not None), 0) - 1

    idx_tail = len(paras)
    for i in range(max(idx_intro + 1, 0), len(paras)):
        t = norm_inline(paras[i].text)
        if paras[i].level is None and (t in ('證據或文件清單：', '證據或文件清單:') or t == '謹狀'
                                       or paras[i].style == '書狀_謹狀'):
            idx_tail = i
            break

    header = paras[:max(idx_intro, 0)]
    intro = paras[idx_intro] if idx_intro >= 0 else None
    body = paras[idx_intro + 1:idx_tail]
    tail = paras[idx_tail:]

    # ---- 狀首 ----
    prev_kind = None
    title_done = False
    for p in header:
        if p.style.startswith('__'):
            continue
        raw = p.text
        t = norm_inline(raw)
        if not t:
            continue
        if not title_done:
            lines.append(f'# {t}')
            title_done = True
            prev_kind = 'title'
            continue
        if prev_kind == 'title' and DATE_ONLY.match(t):
            lines.append(re.sub(r'^(中華民國|民國)\s*', '', t))
            prev_kind = 'date'
            continue
        is_party = p.style == '書狀_狀首當事人' or '\t' in raw
        if is_party:
            text = norm_party(raw)
            role = party_role(text)
            new_group = prev_kind != 'party' or not REPRESENTATIVE.search(role)
            if new_group:
                lines.append('')
            lines.append(text)
            prev_kind = 'party'
        else:
            if prev_kind != 'misc':
                lines.append('')
            lines.append(t)
            prev_kind = 'misc'

    if intro is not None:
        lines.append('')
        lines.append(norm_inline(intro.text))

    # ---- 正文 ----
    counters = [0, 0, 0, 0]
    section = None
    items = []  # (kind, level, text)
    for p in body:
        if p.style == '__list_start__':
            counters = [0, 0, 0, 0]
            continue
        if p.style == '__start_value__':
            if p.level is not None and 0 <= p.level < 4:
                counters[p.level] = p.start_value - 1
                for k in range(p.level + 1, 4):
                    counters[k] = 0
            continue
        if p.style == '__table__':
            items.append(('body', None, p.text))
            warnings.append('文件中有表格，未轉換，請人工補入。')
            continue
        t = norm_inline(p.text)
        if not t:
            continue
        if p.level is None and is_block_heading(t):
            core = t.replace('*', '').replace('　', '').replace(' ', '')
            section = '聲明' if core in DECL_WORDS else '理由'
            counters = [0, 0, 0, 0]
            items.append(('block', None, f'**　　{core}**'))
            continue
        if p.level is None and MAJOR_HEADING.match(t):
            counters = [0, 0, 0, 0]
            items.append(('body', None, t))
            continue
        if p.level is not None and p.numbered:
            lv = p.level
            counters[lv] += 1
            for k in range(lv + 1, 4):
                counters[k] = 0
            items.append(('level', lv, number_prefix(lv, counters[lv]) + t, section))
            continue
        if p.style == '書狀_清單':
            items.append(('body', None, f'> {t}'))
            continue
        items.append(('body', None, t))

    for i, it in enumerate(items):
        kind = it[0]
        text = it[2]
        if kind == 'level' and it[1] == 0 and it[3] != '聲明':
            nxt = items[i + 1] if i + 1 < len(items) else None
            has_child = nxt is not None and (
                (nxt[0] == 'level' and nxt[1] > 0) or nxt[0] == 'body')
            if has_child:
                text = f'**{text}**'
        lines.append('')
        lines.append(text)

    # ---- 狀尾 ----
    in_list = False
    after_gongjian = False
    for p in tail:
        if p.style.startswith('__'):
            continue
        t = norm_inline(p.text)
        if not t:
            continue
        if t.startswith('證據或文件清單'):
            lines.append('')
            lines.append(t)
            in_list = True
            continue
        if in_list and (EVIDENCE_LINE.match(t) or t.startswith('附件')) and t != '謹狀':
            lines.append(t)
            continue
        in_list = False
        if t == '謹狀' or p.style == '書狀_謹狀':
            lines.append('')
            lines.append('謹狀')
            continue
        if '公鑒' in t or '轉呈' in t:
            lines.append(t)
            after_gongjian = True
            continue
        if ROC_DATE.match(t):
            lines.append(t)
            lines.append('')
            continue
        lines.append(t)

    # 收尾：壓縮連續空行
    out = []
    for ln in lines:
        if ln == '' and (not out or out[-1] == ''):
            continue
        out.append(ln)
    while out and out[-1] == '':
        out.pop()
    return '\n'.join(out) + '\n', warnings


# ----------------------------------------------------------------------------
# 差異比對
# ----------------------------------------------------------------------------

def diff_lines(old_text, new_text, old_name, new_name):
    a = [ln.rstrip() for ln in old_text.splitlines() if ln.strip()]
    b = [ln.rstrip() for ln in new_text.splitlines() if ln.strip()]
    return list(difflib.unified_diff(a, b, fromfile=old_name, tofile=new_name, lineterm='', n=0))


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description='將書狀 ODT／DOCX 轉回 draft-pleading 格式的 md 草稿')
    ap.add_argument('input', help='書狀 .odt 或 .docx')
    ap.add_argument('--output', help='輸出 md 路徑（預設：同資料夾、同檔名 .md）')
    ap.add_argument('--diff', help='轉換後與此 md 比對差異（忽略空白行）')
    ap.add_argument('--force', action='store_true', help='允許覆蓋既有的輸出檔')
    args = ap.parse_args()

    src = os.path.abspath(args.input)
    if not os.path.isfile(src):
        print(f'[ERROR] 找不到檔案：{src}', file=sys.stderr)
        return 1
    ext = os.path.splitext(src)[1].lower()
    if ext == '.odt':
        paras = read_odt(src)
    elif ext == '.docx':
        paras = read_docx(src)
    else:
        print('[ERROR] 只支援 .odt 或 .docx', file=sys.stderr)
        return 1

    md, warnings = convert(paras)

    out = os.path.abspath(args.output) if args.output else os.path.splitext(src)[0] + '.md'
    if os.path.abspath(out) == src:
        print('[ERROR] 輸出路徑不可與來源相同', file=sys.stderr)
        return 1
    if os.path.exists(out) and not args.force:
        print(f'[ERROR] 輸出檔已存在，未覆蓋：{out}\n        請改用 --output 指定其他路徑，或確認後加 --force。', file=sys.stderr)
        return 2
    with open(out, 'w', encoding='utf-8', newline='\n') as f:
        f.write(md)

    n_lines = sum(1 for ln in md.splitlines() if ln.strip())
    print(f'[OK] 已轉出：{out}（{n_lines} 行非空白）')
    for w in warnings:
        print(f'[WARN] {w}')

    if args.diff:
        with open(args.diff, encoding='utf-8-sig') as f:
            old = f.read()
        d = diff_lines(old, md, os.path.basename(args.diff), os.path.basename(out))
        if not d:
            print('[DIFF] 與舊 md 一致（忽略空白行）。')
        else:
            changed = sum(1 for ln in d if ln[:1] in '+-' and not ln.startswith(('+++', '---')))
            print(f'[DIFF] 與舊 md 有差異（{changed} 行）：')
            for ln in d:
                print(ln)
    return 0


if __name__ == '__main__':
    sys.exit(main())
