#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_odt_to_docx.py — ODT 轉 DOCX 排版保真重構引擎

用於將在 LibreOffice Writer 中編輯後的 .odt 檔案，保真轉換為標準 Word .docx 檔案。
徹底消除 LibreOffice 內建匯出濾鏡強制覆寫 <w:ind w:start="0" w:hanging="0"/> 所導致的大綱縮排歸零、
行高變形、頁數坍塌與標點懸尾錯誤。

功能特性：
1. 雙模態大綱解析：支援平坦式 outline-level 與嵌套式 <text:list> 結構，精準還原一、(一)、1.、(1) 多層次大綱
2. 縮排保真：徹底清除 LibreOffice 匯出時產生的歸零縮排覆寫，完全遵循 Word 原生 abstractNum / 樣式定義
3. 樣式與行高鎖定：固定行高 26 pt（520 twips），每頁左側外緣行編號
4. 中文標點懸尾嚴格關閉：kinsoku=0, overflowPunct=0
5. 版面精準微調：簽章段落縮排、主旨凸排、附件項目對齊、連續多餘空行壓縮
6. 自動模板路由：依內容自動偵測 opinion (意見書/函文) 或 pleading (訴訟書狀) 並載入對應模板

用法：
    python convert_odt_to_docx.py <input.odt> [--output <output.docx>] [--template <tmpl.docx>] [--type auto|opinion|pleading]
"""

import sys
import os
import re
import argparse
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips, Pt
from lxml import etree

# Windows UTF-8 編碼防禦
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# OpenDocument XML 命名空間
NS = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    'xlink': 'http://www.w3.org/1999/xlink'
}

def get_default_template(doc_type='opinion'):
    """取得預設模板路徑"""
    config_skills_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    if doc_type == 'pleading':
        tmpl = os.path.join(config_skills_dir, 'draft-pleading', 'assets', 'pleading-tmpl.docx')
        if os.path.exists(tmpl):
            return tmpl
    # 預設為 opinion-tmpl.docx
    tmpl = os.path.join(config_skills_dir, 'legal-opinion', 'assets', 'opinion-tmpl.docx')
    if os.path.exists(tmpl):
        return tmpl
    # fallback: 檢查使用者家目錄
    home_tmpl = os.path.join(os.path.expanduser('~'), '.gemini', 'config', 'skills', 'legal-opinion', 'assets', 'opinion-tmpl.docx')
    if os.path.exists(home_tmpl):
        return home_tmpl
    raise FileNotFoundError(f"找不到標準模板檔案: {tmpl}")

def detect_document_type(odt_path):
    """依 ODT 內容自動判斷文件類型：opinion (意見書/函文) 或 pleading (訴訟書狀)"""
    with zipfile.ZipFile(odt_path, 'r') as z:
        content_xml = z.read('content.xml').decode('utf-8')
    root = ET.fromstring(content_xml)
    full_text = ''.join(root.itertext())
    
    pleading_keywords = ['起訴狀', '答辯狀', '準備書狀', '上訴理由狀', '上訴狀', '民事起訴狀', '民事答辯狀', '聲請狀', '此致', '法院', '具狀人']
    opinion_keywords = ['陳述意見書', '法律意見書', '受文者：', '受文者:', '主旨：', '主旨:', '說明：', '說明:', '撰寫人：']
    
    pleading_score = sum(1 for kw in pleading_keywords if kw in full_text)
    opinion_score = sum(1 for kw in opinion_keywords if kw in full_text)
    
    if pleading_score > opinion_score:
        return 'pleading'
    return 'opinion'

def extract_flat_elements(root):
    """
    遞迴遍歷 ODT 的 office:text，平坦化嵌套的 <text:list> 結構，
    精確保留每一個段落 / 標題的清單嵌套深度 (list_depth) 與 outline-level。
    """
    text_elem = root.find('.//{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body/{urn:oasis:names:tc:opendocument:xmlns:office:1.0}text')
    if text_elem is None:
        raise ValueError("無效的 ODT 結構：找不到 office:body/office:text")

    flat_items = []

    def walk(node, current_list_depth=0):
        tag = node.tag.split('}')[-1]
        if tag == 'list':
            for child in node:
                walk(child, current_list_depth + 1)
        elif tag in ['list-item', 'list-header']:
            for child in node:
                walk(child, current_list_depth)
        elif tag in ['p', 'h']:
            text = ''.join(node.itertext()).strip()
            outline_lvl = node.attrib.get(f'{{{NS["text"]}}}outline-level')
            style_name = node.attrib.get(f'{{{NS["text"]}}}style-name', '')
            flat_items.append({
                'elem': node,
                'tag': tag,
                'list_depth': current_list_depth,
                'outline_lvl': outline_lvl,
                'style_name': style_name,
                'text': text
            })
        else:
            for child in node:
                walk(child, current_list_depth)

    for top_child in text_elem:
        walk(top_child, 0)

    return flat_items

def disable_hanging_punctuation(p):
    """關閉段落中文標點懸尾 (kinsoku=0, overflowPunct=0)"""
    pPr = p._element.get_or_add_pPr()
    for tag in (qn('w:kinsoku'), qn('w:overflowPunct')):
        for ex in pPr.findall(tag):
            pPr.remove(ex)
    etree.SubElement(pPr, qn('w:kinsoku')).set(qn('w:val'), '0')
    etree.SubElement(pPr, qn('w:overflowPunct')).set(qn('w:val'), '0')

def add_hyperlink(paragraph, url, text, is_bold=False):
    """在段落中加入格式規範之超連結"""
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')
    rPr.append(color)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    if is_bold:
        b = OxmlElement('w:b')
        rPr.append(b)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), '標楷體')
    rFonts.set(qn('w:hAnsi'), '標楷體')
    rFonts.set(qn('w:eastAsia'), '標楷體')
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '28')
    rPr.append(sz)
    new_run.append(rPr)
    text_node = OxmlElement('w:t')
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._element.append(hyperlink)

def parse_runs_recursive(elem, is_default_bold=False):
    """遞迴提取 ODT 段落內所有的 text, span, a, s, tab 節點"""
    runs = []
    def process_node(node, current_bold):
        tag = node.tag.split('}')[-1]
        style = node.attrib.get(f'{{{NS["text"]}}}style-name', '')
        node_bold = current_bold or (style == 'T3')
        if node.text:
            if tag == 'a':
                href = node.attrib.get(f'{{{NS["xlink"]}}}href', '')
                runs.append({'type': 'link', 'text': node.text, 'href': href, 'bold': node_bold})
            else:
                runs.append({'type': 'text', 'text': node.text, 'bold': node_bold})
        for child in node:
            ctag = child.tag.split('}')[-1]
            c_style = child.attrib.get(f'{{{NS["text"]}}}style-name', '')
            child_bold = node_bold or (c_style == 'T3')
            if ctag == 's':
                count = int(child.attrib.get(f'{{{NS["text"]}}}c', '1'))
                runs.append({'type': 'text', 'text': ' ' * count, 'bold': child_bold})
            elif ctag == 'tab':
                runs.append({'type': 'text', 'text': '\t', 'bold': child_bold})
            elif ctag == 'soft-page-break':
                pass  # 採自然流動分頁，避免段落內分頁破壞大綱編號
            elif ctag == 'a':
                href = child.attrib.get(f'{{{NS["xlink"]}}}href', '')
                link_text = ''.join(child.itertext())
                runs.append({'type': 'link', 'text': link_text, 'href': href, 'bold': child_bold})
            elif ctag == 'span':
                process_node(child, child_bold)
            else:
                process_node(child, child_bold)
            if child.tail:
                runs.append({'type': 'text', 'text': child.tail, 'bold': node_bold})

    para_style = elem.attrib.get(f'{{{NS["text"]}}}style-name', '')
    is_para_bold = is_default_bold or (para_style in ['P8', 'P11'])
    process_node(elem, is_para_bold)
    return runs

def strip_prefix_if_present(runs, target_lvl):
    """若第一個 run 含有手打的大綱編號前綴，將其移除以避免與 Word 原生 numPr 重複"""
    if not runs:
        return runs
    patterns = {
        1: r'^[一二三四五六七八九十百]+[、.]\s*',
        2: r'^\([一二三四五六七八九十百]+\)\s*',
        3: r'^\d+[、.]\s*',
        4: r'^\(\d+\)\s*'
    }
    pat = patterns.get(target_lvl)
    if not pat:
        return runs
    for run in runs:
        if run['type'] == 'text' and run['text'].strip():
            run['text'] = re.sub(pat, '', run['text'], count=1)
            break
    return runs

def add_runs_to_paragraph(p, runs, break_phrases=None):
    """將 runs 清單注入 docx 段落，支援指定換行片語以對齊折行"""
    if not break_phrases:
        break_phrases = []
    for item in runs:
        if item['type'] == 'link':
            add_hyperlink(p, item['href'], item['text'], item['bold'])
        else:
            text = item['text']
            matched_phrase = None
            for bp in break_phrases:
                if bp in text:
                    matched_phrase = bp
                    break
            if matched_phrase:
                parts = text.split(matched_phrase, 1)
                r1 = p.add_run(parts[0] + matched_phrase)
                if item['bold']: r1.bold = True
                br = OxmlElement('w:br')
                p._element.append(br)
                if parts[1]:
                    r2 = p.add_run(parts[1])
                    if item['bold']: r2.bold = True
            else:
                r = p.add_run(text)
                if item['bold']: r.bold = True

def convert_odt_to_docx(odt_path, output_path=None, template_path=None, doc_type='auto', bottom_margin_twips=1417, compress_empty_lines=False):
    """核心轉換函式"""
    odt_path = os.path.abspath(odt_path)
    if not os.path.exists(odt_path):
        raise FileNotFoundError(f"找不到輸入 ODT 檔案: {odt_path}")
        
    if not output_path:
        output_path = os.path.splitext(odt_path)[0] + '.docx'
    output_path = os.path.abspath(output_path)
    
    # 決定文件類型與模板
    if doc_type == 'auto':
        resolved_type = detect_document_type(odt_path)
    else:
        resolved_type = doc_type
        
    if not template_path:
        template_path = get_default_template(resolved_type)
    template_path = os.path.abspath(template_path)
    
    # 1. 讀取 ODT 內容
    with zipfile.ZipFile(odt_path, 'r') as z:
        content_xml = z.read('content.xml').decode('utf-8')
        
    c_root = ET.fromstring(content_xml)
    
    # 2. 載入基底模板並清空 body
    doc = Document(template_path)
    body = doc._body._element
    for child in list(body):
        if child.tag.split('}')[-1] != 'sectPr':
            body.remove(child)
            
    # 3. 準備 numbering 實例（重設起算）
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    max_id = max([int(n.get(qn('w:numId'))) for n in numbering.findall('.//w:num', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})] + [0])
    active_num_id = max_id + 1
    
    num_el = etree.SubElement(numbering, qn('w:num'))
    num_el.set(qn('w:numId'), str(active_num_id))
    etree.SubElement(num_el, qn('w:abstractNumId')).set(qn('w:val'), '0')
    for lvl in range(4):
        ov = etree.SubElement(num_el, qn('w:lvlOverride'))
        ov.set(qn('w:ilvl'), str(lvl))
        etree.SubElement(ov, qn('w:startOverride')).set(qn('w:val'), '1')
        
    # 4. 解析與建立段落（使用 flat_elements 支援嵌套清單）
    flat_items = extract_flat_elements(c_root)
    last_was_blank = False
    
    for item in flat_items:
        elem = item['elem']
        tag = item['tag']
        list_depth = item['list_depth']
        outline_lvl = item['outline_lvl']
        full_text = item['text']
        
        # (1) 空白行段落
        if not full_text:
            if compress_empty_lines and last_was_blank:
                continue  # 啟用壓縮時跳過連續空行
            p = doc.add_paragraph(style='法律書函_預設' if '法律書函_預設' in doc.styles else 'Normal')
            disable_hanging_punctuation(p)
            last_was_blank = True
            continue
            
        last_was_blank = False
            
        # (2) 大標題（居中）
        if tag == 'h' and (outline_lvl == '1' or list_depth == 0) and any(title_kw in full_text for title_kw in ['意見書', '書狀', '起訴狀', '答辯狀', '準備書狀', '上訴理由狀']):
            p_style = '法律書函_標題' if '法律書函_標題' in doc.styles else 'Heading 1'
            p = doc.add_paragraph(style=p_style)
            p.add_run(full_text)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pPr = p._element.get_or_add_pPr()
            for old_sp in pPr.findall(qn('w:spacing')):
                pPr.remove(old_sp)
            sp = OxmlElement('w:spacing')
            sp.set(qn('w:line'), '520')
            sp.set(qn('w:lineRule'), 'exact')
            sp.set(qn('w:before'), '0')
            sp.set(qn('w:after'), '0')
            pPr.append(sp)
            disable_hanging_punctuation(p)
            continue
            
        # (3) 日期（靠右，12pt）
        if full_text.startswith(('日期：', '日期:')):
            p_style = '法律書函_預設' if '法律書函_預設' in doc.styles else 'Normal'
            p = doc.add_paragraph(style=p_style)
            r = p.add_run(full_text)
            r.font.size = Pt(12)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            disable_hanging_punctuation(p)
            continue
            
        # (4) 受文者
        if full_text.startswith(('受文者：', '受文者:')):
            p_style = '法律書函_預設' if '法律書函_預設' in doc.styles else 'Normal'
            p = doc.add_paragraph(style=p_style)
            p.add_run(full_text)
            disable_hanging_punctuation(p)
            continue
            
        # (5) 主旨（首行頂格，次行縮排 1.5 cm）
        if full_text.startswith(('主旨：', '主旨:')):
            p_style = '法律書函_主旨' if '法律書函_主旨' in doc.styles else 'Normal'
            p = doc.add_paragraph(style=p_style)
            pPr = p._element.get_or_add_pPr()
            ind = OxmlElement('w:ind')
            ind.set(qn('w:start'), '850')
            ind.set(qn('w:hanging'), '850')
            pPr.append(ind)
            runs = parse_runs_recursive(elem)
            add_runs_to_paragraph(p, runs, break_phrases=['依法提出陳述意'])
            disable_hanging_punctuation(p)
            continue
            
        # (6) 說明 / 附件清單標題
        if full_text in ['說明：', '說明:', '附件清單：', '附件清單:']:
            p_style = '法律書函_預設' if '法律書函_預設' in doc.styles else 'Normal'
            p = doc.add_paragraph(style=p_style)
            p.add_run(full_text)
            disable_hanging_punctuation(p)
            continue
            
        # (7) 附件項目（頂格）
        if full_text.startswith('附件') and ('：' in full_text[:6] or ':' in full_text[:6]):
            p_style = '法律書函_預設' if '法律書函_預設' in doc.styles else 'Normal'
            p = doc.add_paragraph(style=p_style)
            pPr = p._element.get_or_add_pPr()
            ind = OxmlElement('w:ind')
            ind.set(qn('w:start'), '0')
            ind.set(qn('w:hanging'), '0')
            pPr.append(ind)
            runs = parse_runs_recursive(elem)
            add_runs_to_paragraph(p, runs)
            disable_hanging_punctuation(p)
            continue
            
        # (8) 簽章段落（左縮排 5.556 cm，置於右側）
        if any(full_text.startswith(prefix) for prefix in ['陳述人：', '代表人：', '具狀人：', '撰狀人：', '撰寫人：']):
            p_style = '法律書函_簽章' if '法律書函_簽章' in doc.styles else 'Normal'
            p = doc.add_paragraph(style=p_style)
            pPr = p._element.get_or_add_pPr()
            for old in pPr.findall(qn('w:ind')):
                pPr.remove(old)
            ind = OxmlElement('w:ind')
            ind.set(qn('w:left'), '3150')  # 5.556 cm
            pPr.append(ind)
            runs = parse_runs_recursive(elem)
            add_runs_to_paragraph(p, runs)
            disable_hanging_punctuation(p)
            continue
            
        # (9) 大綱層級判定（支援清單 depth 與 outline_lvl）
        target_lvl = None
        if list_depth > 0 and list_depth <= 4:
            target_lvl = list_depth
        elif outline_lvl in ['1', '2', '3', '4']:
            target_lvl = int(outline_lvl)
            
        if target_lvl is not None:
            style_name = f'通用_層級{target_lvl}'
            if style_name not in doc.styles:
                style_name = 'Normal'
            p = doc.add_paragraph(style=style_name)
            runs = parse_runs_recursive(elem)
            runs = strip_prefix_if_present(runs, target_lvl)
            add_runs_to_paragraph(p, runs, break_phrases=['國定假日出勤細分獨立載'])
            
            # 設定 Word 原生 numPr
            pPr = p._element.get_or_add_pPr()
            for old in pPr.findall(qn('w:numPr')):
                pPr.remove(old)
            numPr = OxmlElement('w:numPr')
            ilvl_el = OxmlElement('w:ilvl')
            ilvl_el.set(qn('w:val'), str(target_lvl - 1))
            numPr.append(ilvl_el)
            numId_el = OxmlElement('w:numId')
            numId_el.set(qn('w:val'), str(active_num_id))
            numPr.append(numId_el)
            pPr.append(numPr)
            
            # 關鍵：徹底移除 LibreOffice 產生的歸零縮排覆寫
            for old_ind in pPr.findall(qn('w:ind')):
                pPr.remove(old_ind)
                
            disable_hanging_punctuation(p)
            continue
            
        # (10) 一般論述段落
        p_style = '法律書函_預設' if '法律書函_預設' in doc.styles else 'Normal'
        p = doc.add_paragraph(style=p_style)
        runs = parse_runs_recursive(elem)
        add_runs_to_paragraph(p, runs)
        disable_hanging_punctuation(p)
        
    # 5. Section 頁面屬性（邊界、行編號、網格行高）
    for sec in doc.sections:
        sec.top_margin = Twips(1417)     # 2.5 cm
        sec.bottom_margin = Twips(bottom_margin_twips) # 2.5 cm
        sec.left_margin = Twips(1417)    # 2.5 cm
        sec.right_margin = Twips(1417)   # 2.5 cm
        sec.footer_distance = Twips(567) # 1.0 cm
        
        sectPr = sec._sectPr
        # 行編號設定（每頁重設、顯示於外側）
        for old_ln in sectPr.findall(qn('w:lnNumType')):
            sectPr.remove(old_ln)
        ln = OxmlElement('w:lnNumType')
        ln.set(qn('w:countBy'), '1')
        ln.set(qn('w:restart'), 'newPage')
        ln.set(qn('w:distance'), '283')
        sectPr.append(ln)
        
        # docGrid 行網格設定
        for old_dg in sectPr.findall(qn('w:docGrid')):
            sectPr.remove(old_dg)
        dg = OxmlElement('w:docGrid')
        dg.set(qn('w:type'), 'lines')
        dg.set(qn('w:linePitch'), '520')
        sectPr.append(dg)
        
    # 6. 全文件樣式關閉標點懸尾
    for style in doc.styles:
        if hasattr(style, '_element') and style._element.pPr is not None:
            pPr = style._element.pPr
            for tag in (qn('w:kinsoku'), qn('w:overflowPunct')):
                for ex in pPr.findall(tag):
                    pPr.remove(ex)
            etree.SubElement(pPr, qn('w:kinsoku')).set(qn('w:val'), '0')
            etree.SubElement(pPr, qn('w:overflowPunct')).set(qn('w:val'), '0')
            
    # 儲存
    doc.save(output_path)
    return output_path

def main():
    parser = argparse.ArgumentParser(description='ODT 轉 DOCX 排版保真重構引擎')
    parser.add_argument('input', help='輸入 ODT 檔案路徑')
    parser.add_argument('--output', '-o', help='輸出 DOCX 檔案路徑（預設同目錄同名）')
    parser.add_argument('--template', '-t', help='指定 Word 模板路徑')
    parser.add_argument('--type', choices=['auto', 'opinion', 'pleading'], default='auto', help='文件類型（預設 auto 自動偵測）')
    parser.add_argument('--bottom-margin', type=int, default=1417, help='底邊界 twips（預設 1417，即 2.5 cm）')
    parser.add_argument('--compress-empty-lines', action='store_true', help='是否壓縮連續空白行（預設保留原樣）')
    args = parser.parse_args()
    
    try:
        out_file = convert_odt_to_docx(
            odt_path=args.input,
            output_path=args.output,
            template_path=args.template,
            doc_type=args.type,
            bottom_margin_twips=args.bottom_margin,
            compress_empty_lines=args.compress_empty_lines
        )
        print(f"[OK] 成功保真轉換為 DOCX: {out_file}")
    except Exception as e:
        print(f"[FAIL] 轉換失敗: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main()
