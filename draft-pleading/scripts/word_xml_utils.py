# word_xml_utils.py

"""
專門處理 python-docx 尚不支援的深層 XML 操作與 Numbering 定義檔竄改。
"""

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

NSMAP = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ANCHOR_STYLE = '通用_層級1'

def find_and_remove_anchor(doc):
    """
    錨點反查：找到套用 ANCHOR_STYLE 的段落，取得其 numId，
    再從 numbering 反查 abstractNumId，最後刪除錨點段落。

    回傳 (num_id, abstract_num_id)
    """
    anchor_para = None
    anchor_idx = None

    for i, p in enumerate(doc.paragraphs):
        if p.style.name == ANCHOR_STYLE:
            anchor_para = p
            anchor_idx = i
            break

    if anchor_para is None:
        raise RuntimeError(f"模板中找不到套用 '{ANCHOR_STYLE}' 樣式的錨點段落")

    # 從段落 XML 取得 numId
    pPr = anchor_para._element.pPr
    if pPr is None or pPr.numPr is None or pPr.numPr.numId is None:
        # 嘗試從 basedOn 鏈反查
        num_id = _trace_num_id_from_style(doc, anchor_para.style)
    else:
        num_id = pPr.numPr.numId.val

    # 從 numbering 反查 abstractNumId
    abstract_num_id = _get_abstract_num_id(doc, num_id)

    # 刪除錨點段落
    anchor_para._element.getparent().remove(anchor_para._element)
    print(f"[INFO] 錨點反查成功: numId={num_id}, abstractNumId={abstract_num_id}")
    print(f"[INFO] 已刪除錨點段落 (index={anchor_idx})")

    return num_id, abstract_num_id


def _trace_num_id_from_style(doc, style):
    """沿 basedOn 鏈追蹤 numId。"""
    curr = style
    while curr:
        pPr = curr._element.find('.//w:pPr', NSMAP)
        if pPr is not None:
            numPr = pPr.find('w:numPr', NSMAP)
            if numPr is not None:
                nid = numPr.find('w:numId', NSMAP)
                if nid is not None:
                    return int(nid.get(qn('w:val')))
        if curr.base_style:
            curr = curr.base_style
        else:
            break
    raise RuntimeError("無法從樣式鏈取得編號 ID")


def _get_abstract_num_id(doc, num_id):
    """從 numbering.xml 中取得 numId 對應的 abstractNumId。"""
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    for num_el in numbering.findall('.//w:num', NSMAP):
        nid = int(num_el.get(qn('w:numId')))
        if nid == num_id:
            anid_el = num_el.find('w:abstractNumId', NSMAP)
            if anid_el is not None:
                return int(anid_el.get(qn('w:val')))
    raise RuntimeError(f"numbering.xml 中找不到 numId={num_id}")


def clear_body(doc):
    """清空模板的 body 段落內容，保留 sectPr（頁面設定）。"""
    body = doc.element.body
    # 保留 sectPr
    sect_pr = body.find('w:sectPr', NSMAP)

    # 移除所有段落和表格
    for child in list(body):
        if child.tag != qn('w:sectPr'):
            body.remove(child)

    print("[INFO] 已清空模板段落內容")


def enable_line_numbering(doc):
    """為文件所有節 (Section) 啟用行編號，每頁重新起算。"""
    for section in doc.sections:
        sectPr = section._sectPr
        # 尋找是否已有 w:lnNumType 標籤
        lnNumType = sectPr.find(qn('w:lnNumType'))
        if lnNumType is None:
            lnNumType = OxmlElement('w:lnNumType')
            sectPr.append(lnNumType)

        # 設定屬性：每行編號、每頁重新起算
        lnNumType.set(qn('w:countBy'), '1')
        lnNumType.set(qn('w:restart'), 'newPage')


def disable_hanging_punctuation(doc):
    """為文件所有段落與樣式直接關閉「允許標點符號溢出邊界」（懸尾）設定。"""
    # 1. 處理所有樣式
    for style in doc.styles:
        if hasattr(style, '_element') and style._element.pPr is not None:
            pPr = style._element.pPr
            # 移除既有的設定
            for tag in (qn('w:kinsoku'), qn('w:overflowPunct')):
                for ex in pPr.findall(tag):
                    pPr.remove(ex)
            # 強制設為 0
            etree.SubElement(pPr, qn('w:kinsoku')).set(qn('w:val'), '0')
            etree.SubElement(pPr, qn('w:overflowPunct')).set(qn('w:val'), '0')

    # 2. 處理所有段落 (Direct formatting 覆寫)
    for p in doc.paragraphs:
        pPr = p._element.get_or_add_pPr()
        for tag in (qn('w:kinsoku'), qn('w:overflowPunct')):
            for ex in pPr.findall(tag):
                pPr.remove(ex)
        etree.SubElement(pPr, qn('w:kinsoku')).set(qn('w:val'), '0')
        etree.SubElement(pPr, qn('w:overflowPunct')).set(qn('w:val'), '0')


def create_override_num(doc, abstract_num_id):
    """
    動態建立帶有 startOverride=1 的新 w:num 實例。
    確保後續的 通用_層級1 編號從「一、」重新起算。

    回傳新的 numId。
    """
    numbering = doc.part.numbering_part.numbering_definitions._numbering

    # 找到目前最大的 numId
    max_num_id = 0
    for num_el in numbering.findall('.//w:num', NSMAP):
        nid = int(num_el.get(qn('w:numId')))
        if nid > max_num_id:
            max_num_id = nid
    new_num_id = max_num_id + 1

    # 建立新的 w:num 元素
    new_num = etree.SubElement(numbering, qn('w:num'))
    new_num.set(qn('w:numId'), str(new_num_id))

    # 指向同一個 abstractNum
    abs_ref = etree.SubElement(new_num, qn('w:abstractNumId'))
    abs_ref.set(qn('w:val'), str(abstract_num_id))

    # 為所有 4 個層級加入 startOverride=1
    for ilvl in range(4):
        override = etree.SubElement(new_num, qn('w:lvlOverride'))
        override.set(qn('w:ilvl'), str(ilvl))
        start_override = etree.SubElement(override, qn('w:startOverride'))
        start_override.set(qn('w:val'), '1')

    print(f"[INFO] 動態建立 numId={new_num_id} (abstractNumId={abstract_num_id}, startOverride=1)")
    return new_num_id


def create_l2_reset_num(doc, abstract_num_id, current_num_id):
    """
    書狀相容層：當進入新的第一層主段（#### 一、→二、→三、）後，
    建立一個只對 ilvl=1 做 startOverride=1 的新 numId，
    確保第二層 (一)(二)... 在每個新第一層下重新起算。
    """
    numbering = doc.part.numbering_part.numbering_definitions._numbering

    max_num_id = 0
    for num_el in numbering.findall('.//w:num', NSMAP):
        nid = int(num_el.get(qn('w:numId')))
        if nid > max_num_id:
            max_num_id = nid
    new_num_id = max_num_id + 1

    new_num = etree.SubElement(numbering, qn('w:num'))
    new_num.set(qn('w:numId'), str(new_num_id))

    abs_ref = etree.SubElement(new_num, qn('w:abstractNumId'))
    abs_ref.set(qn('w:val'), str(abstract_num_id))

    # 只重設 ilvl=1（第二層 (一)），不干擾第一層的進行中計數
    override = etree.SubElement(new_num, qn('w:lvlOverride'))
    override.set(qn('w:ilvl'), '1')
    start_override = etree.SubElement(override, qn('w:startOverride'))
    start_override.set(qn('w:val'), '1')

    print(f"[INFO] 第二層重新起算: 建立 numId={new_num_id} (ilvl=1 startOverride=1)")
    return new_num_id


def set_num_pr(para, num_id, ilvl, outline_level=None):
    """在段落的 pPr 中設定 numPr（覆寫樣式繼承的值），並可選設定大綱層級。"""
    pPr = para._element.get_or_add_pPr()

    # 1. 設定編號 numPr
    existing_num = pPr.find(qn('w:numPr'))
    if existing_num is not None:
        pPr.remove(existing_num)

    numPr = etree.SubElement(pPr, qn('w:numPr'))
    ilvl_el = etree.SubElement(numPr, qn('w:ilvl'))
    ilvl_el.set(qn('w:val'), str(ilvl))
    numId_el = etree.SubElement(numPr, qn('w:numId'))
    numId_el.set(qn('w:val'), str(num_id))

    # 2. 設定大綱層級 outlineLvl (影響 PDF 索引與 Writer 導航)
    existing_outline = pPr.find(qn('w:outlineLvl'))
    if existing_outline is not None:
        pPr.remove(existing_outline)

    if outline_level is not None:
        # 0 為最上層標題
        outline_el = etree.SubElement(pPr, qn('w:outlineLvl'))
        outline_el.set(qn('w:val'), str(outline_level))
    else:
        # 強制設為 9 (Body Text)，覆蓋樣式繼承的大綱層級，
        # 避免聲明區段落出現在導航面板與 PDF 書籤中
        outline_el = etree.SubElement(pPr, qn('w:outlineLvl'))
        outline_el.set(qn('w:val'), '9')
