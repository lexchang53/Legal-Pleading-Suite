#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
舊 ODT 結構升級修復工具 (Upgrade Existing ODT to support Tab demote/promote)
"""
import argparse
import os
import re
import tempfile
import zipfile
from pathlib import Path

LEVEL_STYLE_MAP = {
    '通用_5f_層級1': 1, '通用_5f_層級2': 2,
    '通用_5f_層級3': 3, '通用_5f_層級4': 4,
    '通用_5f_清單': 5,
    '通用_層級1': 1, '通用_層級2': 2,
    '通用_層級3': 3, '通用_層級4': 4,
    '通用_清單': 5,
    '書狀_5f_層級1': 1, '書狀_5f_層級2': 2,
    '書狀_5f_層級3': 3, '書狀_5f_層級4': 4,
    '書狀_5f_清單': 5,
    '書狀_層級1': 1, '書狀_層級2': 2,
    '書狀_層級3': 3, '書狀_層級4': 4,
    '書狀_清單': 5,
}

OUTLINE_STYLE_XML = (
    '<text:outline-style style:name="Outline" '
    'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
    'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
    'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
    'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">'
    '<text:outline-level-style text:level="1" '
    'loext:num-list-format="%1%、" style:num-suffix="、" style:num-format="一, 二, 三, ...">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="1cm" fo:text-indent="-1cm" fo:margin-left="1cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="2" '
    'loext:num-list-format="(%2%)" style:num-prefix="(" style:num-suffix=")" style:num-format="一, 二, 三, ...">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="1.499cm" fo:text-indent="-1cm" fo:margin-left="1.499cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="3" '
    'loext:num-list-format="%3%." style:num-suffix="." style:num-format="1">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="1.499cm" fo:text-indent="-0.499cm" fo:margin-left="1.499cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="4" '
    'loext:num-list-format="(%4%)" style:num-prefix="(" style:num-suffix=")" style:num-format="1">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="1.998cm" fo:text-indent="-0.499cm" fo:margin-left="1.998cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="5" '
    'loext:num-list-format="%5%、" style:num-suffix="、" style:num-format="甲, 乙, 丙, ...">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="2.499cm" fo:text-indent="-1.05cm" fo:margin-left="2.499cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="6" loext:num-list-format="%6%" style:num-format="">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="nothing"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="7" loext:num-list-format="%7%" style:num-format="">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="nothing"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="8" loext:num-list-format="%8%" style:num-format="">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="nothing"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="9" loext:num-list-format="%9%" style:num-format="">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="nothing"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '</text:outline-style>'
)

def register_all_namespaces(xml_string):
    namespaces = re.findall(r'xmlns:([^=]+)="([^"]+)"', xml_string)
    for prefix, uri in namespaces:
        import xml.etree.ElementTree as ET
        ET.register_namespace(prefix, uri)

def process_content_xml(data):
    try:
        from lxml import etree as _et
        use_lxml = True
    except ImportError:
        use_lxml = False

    if use_lxml:
        root = _et.fromstring(data)
        nsmap = root.nsmap
        def qn(prefix, localname):
            return f"{{{nsmap[prefix]}}}{localname}"

        # 1. 清理 styles 段落樣式屬性
        auto_styles = {}
        for style in root.findall('.//style:style', nsmap):
            name = style.get(qn('style', 'name'))
            parent = style.get(qn('style', 'parent-style-name'))
            if name and parent:
                if parent in LEVEL_STYLE_MAP:
                    auto_styles[name] = parent
                if parent in LEVEL_STYLE_MAP or parent == '通用多層清單':
                    for attr in (qn('style', 'list-style-name'), qn('style', 'list-level')):
                        if attr in style.attrib:
                            del style.attrib[attr]
        
        # 2. 移除 <text:list> 清單外殼
        while True:
            lists = root.findall('.//text:list', nsmap)
            if not lists:
                break
            for lst in lists:
                parent_node = lst.getparent()
                if parent_node is None:
                    continue
                index = parent_node.index(lst)
                for list_item in lst.findall('text:list-item', nsmap):
                    for child in list(list_item):
                        parent_node.insert(index, child)
                        index += 1
                parent_node.remove(lst)

        # 3. 轉換 <text:p> 為 <text:h> 標題
        for p in root.findall('.//text:p', nsmap):
            sname = p.get(qn('text', 'style-name'))
            target_level = None
            if sname in LEVEL_STYLE_MAP:
                target_level = str(LEVEL_STYLE_MAP[sname])
            elif sname in auto_styles and auto_styles[sname] in LEVEL_STYLE_MAP:
                target_level = str(LEVEL_STYLE_MAP[auto_styles[sname]])

            if target_level:
                p.tag = qn('text', 'h')
                p.set(qn('text', 'outline-level'), target_level)

        # 4. 清理書籤
        for bookmark in root.findall('.//text:bookmark-start', nsmap):
            bookmark.getparent().remove(bookmark)
        for bookmark in root.findall('.//text:bookmark-end', nsmap):
            bookmark.getparent().remove(bookmark)
        for bookmark in root.findall('.//text:bookmark', nsmap):
            bookmark.getparent().remove(bookmark)

        # 5. 關閉標點符號懸尾
        for style in root.findall('.//style:style', nsmap):
            if style.attrib.get(qn('style', 'family')) == 'paragraph':
                props = style.find('.//style:paragraph-properties', nsmap)
                if props is None:
                    props = _et.SubElement(style, qn('style', 'paragraph-properties'))
                props.attrib[qn('style', 'punctuation-wrap')] = 'simple'

        return _et.tostring(root, xml_declaration=True, encoding='UTF-8')

    else:
        # Fallback using standard xml.etree.ElementTree
        import xml.etree.ElementTree as ET
        register_all_namespaces(data.decode('utf-8', errors='replace'))
        root = ET.fromstring(data)
        
        ns = {
            'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
            'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
        }

        auto_styles = {}
        for style in root.findall('.//style:style', ns):
            name = style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name')
            parent = style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name')
            if name and parent:
                if parent in LEVEL_STYLE_MAP:
                    auto_styles[name] = parent
                if parent in LEVEL_STYLE_MAP or parent == '通用多層清單':
                    for attr in ('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}list-style-name', 
                                 '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}list-level'):
                        if attr in style.attrib:
                            del style.attrib[attr]

        # 移除清單外殼
        while True:
            lists = root.findall('.//text:list', ns)
            if not lists:
                break
            parent_map = {c: p for p in root.iter() for c in p}
            for lst in lists:
                parent_node = parent_map.get(lst)
                if parent_node is None:
                    continue
                index = list(parent_node).index(lst)
                for list_item in lst.findall('text:list-item', ns):
                    for child in list(list_item):
                        parent_node.insert(index, child)
                        index += 1
                        parent_map[child] = parent_node
                parent_node.remove(lst)
                parent_map.pop(lst, None)

        # 轉換段落標籤為大綱標籤
        for p in root.findall('.//text:p', ns):
            sname = p.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name')
            target_level = None
            if sname in LEVEL_STYLE_MAP:
                target_level = str(LEVEL_STYLE_MAP[sname])
            elif sname in auto_styles and auto_styles[sname] in LEVEL_STYLE_MAP:
                target_level = str(LEVEL_STYLE_MAP[auto_styles[sname]])

            if target_level:
                p.tag = '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}h'
                p.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}outline-level'] = target_level

        # 清理書籤
        for parent in root.iter():
            to_remove = []
            for child in parent:
                tag = child.tag.split('}')[-1]
                if tag in ('bookmark', 'bookmark-start', 'bookmark-end'):
                    to_remove.append(child)
            for child in to_remove:
                parent.remove(child)

        # 關閉懸尾
        for style in root.findall('.//style:style', ns):
            if style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family') == 'paragraph':
                props = style.find('.//style:paragraph-properties', ns)
                if props is None:
                    props = ET.SubElement(style, '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}paragraph-properties')
                props.attrib['{urn:oasis:names:tc:opendocument:xmlns:style:1.0}punctuation-wrap'] = 'simple'

        return ET.tostring(root, encoding='utf-8')

def process_styles_xml(data):
    try:
        from lxml import etree as _et
        use_lxml = True
    except ImportError:
        use_lxml = False

    if use_lxml:
        root = _et.fromstring(data)
        nsmap = root.nsmap
        def qn(prefix, localname):
            return f"{{{nsmap[prefix]}}}{localname}"

                # 1.1. 提取原本 Word 中的縮排設定，重現精準距離
        style_to_list = {}
        for style in root.findall('.//style:style', nsmap):
            name = style.get(qn('style', 'name'))
            if name in LEVEL_STYLE_MAP:
                ls_name = style.get(qn('style', 'list-style-name'))
                if ls_name:
                    style_to_list[name] = ls_name
                    
        extracted_levels = {}
        for ls_name in set(style_to_list.values()):
            for ls_node in root.findall('.//text:list-style', nsmap):
                if ls_node.get(qn('style', 'name')) == ls_name:
                    for lvl_node in ls_node.findall('./text:list-level-style-number', nsmap):
                        lvl_str = lvl_node.get(qn('text', 'level'))
                        if lvl_str:
                            lvl = int(lvl_str)
                            if 1 <= lvl <= 5:
                                props = lvl_node.find('./style:list-level-properties', nsmap)
                                margin_left = None
                                text_indent = None
                                tab_stop = None
                                if props is not None:
                                     margin_left = props.get(qn('fo', 'margin-left'))
                                     text_indent = props.get(qn('fo', 'text-indent'))
                                     label_align = props.find('./style:list-level-label-alignment', nsmap)
                                     if label_align is not None:
                                         tab_stop = label_align.get(qn('text', 'list-tab-stop-position'))
                                         if not margin_left:
                                             margin_left = label_align.get(qn('fo', 'margin-left'))
                                         if not text_indent:
                                             text_indent = label_align.get(qn('fo', 'text-indent'))
                                
                                extracted_levels[lvl] = {
                                     'margin_left': margin_left,
                                     'text_indent': text_indent,
                                     'tab_stop': tab_stop,
                                     'num_format': lvl_node.get(qn('style', 'num-format')),
                                     'num_prefix': lvl_node.get(qn('style', 'num-prefix')),
                                     'num_suffix': lvl_node.get(qn('style', 'num-suffix')),
                                     'display_name': lvl_node.get(qn('text', 'style-name'))
                                }

        # 1.2. 動態建構大綱樣式 XML
        fallback = {
            1: {'margin_left': '1.5cm', 'text_indent': '-1.5cm', 'tab_stop': '1.5cm', 'num_format': '一, 二, 三, ...', 'num_suffix': '、', 'num_prefix': ''},
            2: {'margin_left': '2.25cm', 'text_indent': '-1.5cm', 'tab_stop': '2.25cm', 'num_format': '一, 二, 三, ...', 'num_suffix': '', 'num_prefix': '('},
            3: {'margin_left': '3cm', 'text_indent': '-1cm', 'tab_stop': '3cm', 'num_format': '1', 'num_suffix': '.', 'num_prefix': ''},
            4: {'margin_left': '3.75cm', 'text_indent': '-1.25cm', 'tab_stop': '3.75cm', 'num_format': '1', 'num_suffix': '', 'num_prefix': '('},
            5: {'margin_left': '4.5cm', 'text_indent': '-1.5cm', 'tab_stop': '4.5cm', 'num_format': '甲, 乙, 丙, ...', 'num_suffix': '、', 'num_prefix': ''}
        }
        for l in range(6, 11):
            fallback[l] = {'margin_left': '0cm', 'text_indent': '0cm', 'tab_stop': '', 'num_format': '', 'num_suffix': '', 'num_prefix': ''}
         
        xml_parts = [
            '<text:outline-style style:name="Outline" '
            'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
            'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
            'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
            'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">'
        ]
        for lvl in range(1, 11):
            val = extracted_levels.get(lvl, fallback.get(lvl))
            margin_left = val.get('margin_left') or fallback[lvl]['margin_left']
            text_indent = val.get('text_indent') or fallback[lvl]['text_indent']
            tab_stop = val.get('tab_stop') or fallback[lvl]['tab_stop']
            num_format = val.get('num_format') if val.get('num_format') is not None else fallback[lvl]['num_format']
            num_prefix = val.get('num_prefix') or fallback[lvl]['num_prefix'] or ''
            num_suffix = val.get('num_suffix') or fallback[lvl]['num_suffix'] or ''
            display_name = val.get('display_name')
             
            if num_prefix or num_suffix:
                num_list_format = f"{num_prefix}%{lvl}%{num_suffix}"
            else:
                if lvl == 1:
                    num_list_format = "%1%、"
                elif lvl == 3:
                    num_list_format = "%3%."
                else:
                    num_list_format = f"%{lvl}%"
                     
            style_name_attr = f' text:style-name="{display_name}"' if display_name else ''
            prefix_attr = f' style:num-prefix="{num_prefix}"' if num_prefix else ''
            suffix_attr = f' style:num-suffix="{num_suffix}"' if num_suffix else ''
             
            xml_parts.append(
                f'<text:outline-level-style text:level="{lvl}"{style_name_attr} '
                f'loext:num-list-format="{num_list_format}"{prefix_attr}{suffix_attr} style:num-format="{num_format}">'
            )
            xml_parts.append('<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">')
            tab_stop_attr = f' text:list-tab-stop-position="{tab_stop}" text:label-followed-by="listtab"' if tab_stop else ' text:label-followed-by="nothing"'
            xml_parts.append(
                f'<style:list-level-label-alignment{tab_stop_attr} fo:text-indent="{text_indent}" fo:margin-left="{margin_left}"/>'
            )
            xml_parts.append('</style:list-level-properties>')
            xml_parts.append('</text:outline-level-style>')
         
        xml_parts.append('</text:outline-style>')
        dynamic_outline_xml = "".join(xml_parts)

        # 1. 重設大綱樣式
        for style in root.findall('.//style:style', nsmap):
            name = style.get(qn('style', 'name'))
            if name in LEVEL_STYLE_MAP:
                level = str(LEVEL_STYLE_MAP[name])
                if qn('style', 'list-style-name') in style.attrib:
                    del style.attrib[qn('style', 'list-style-name')]
                if qn('style', 'list-level') in style.attrib:
                    del style.attrib[qn('style', 'list-level')]
                style.set(qn('style', 'default-outline-level'), level)
            elif name == '通用多層清單':
                if qn('style', 'list-style-name') in style.attrib:
                    del style.attrib[qn('style', 'list-style-name')]
                if qn('style', 'list-level') in style.attrib:
                    del style.attrib[qn('style', 'list-level')]
                style.set(qn('style', 'default-outline-level'), '5')

        # 2. 更新或注入 Outline 大綱樣式
        working_outline_style = _et.fromstring(dynamic_outline_xml)
        outline_style = root.find('.//text:outline-style', nsmap)
        if outline_style is not None:
            parent_node = outline_style.getparent()
            idx = parent_node.index(outline_style)
            parent_node.remove(outline_style)
            parent_node.insert(idx, working_outline_style)
        else:
            styles_node = root.find('.//office:styles', nsmap)
            if styles_node is not None:
                styles_node.append(working_outline_style)

        # 3. 關閉標點符號懸尾
        for style in root.findall('.//style:style', nsmap):
            if style.attrib.get(qn('style', 'family')) == 'paragraph':
                props = style.find('.//style:paragraph-properties', nsmap)
                if props is None:
                    props = _et.SubElement(style, qn('style', 'paragraph-properties'))
                props.attrib[qn('style', 'punctuation-wrap')] = 'simple'

        # 4. 行號改為外側
        ln_config = root.find('.//text:linenumbering-configuration', nsmap)
        if ln_config is not None:
            ln_config.attrib[qn('text', 'number-lines')] = 'true'
            ln_config.attrib[qn('text', 'number-position')] = 'outside'
            ln_config.attrib[qn('style', 'num-format')] = '1'
            ln_config.attrib[qn('text', 'increment')] = '1'
        else:
            office_styles = root.find('.//office:styles', nsmap)
            if office_styles is not None:
                ln_config = _et.SubElement(office_styles, qn('text', 'linenumbering-configuration'))
                ln_config.attrib[qn('text', 'number-lines')] = 'true'
                ln_config.attrib[qn('text', 'number-position')] = 'outside'
                ln_config.attrib[qn('style', 'num-format')] = '1'
                ln_config.attrib[qn('text', 'increment')] = '1'
                ln_config.attrib[qn('text', 'offset')] = '1cm'
                ln_config.attrib[qn('text', 'count-empty-lines')] = 'true'
                ln_config.attrib[qn('text', 'count-in-floating-frames')] = 'false'
                ln_config.attrib[qn('text', 'restart-on-page')] = 'true'

        return _et.tostring(root, xml_declaration=True, encoding='UTF-8')

    else:
        # Fallback using standard xml.etree.ElementTree
        import xml.etree.ElementTree as ET
        register_all_namespaces(data.decode('utf-8', errors='replace'))
        root = ET.fromstring(data)
        ns = {
            'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
            'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
        }

        for style in root.findall('.//style:style', ns):
            name = style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name')
            if name in LEVEL_STYLE_MAP:
                level = str(LEVEL_STYLE_MAP[name])
                style.attrib.pop('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}list-style-name', None)
                style.attrib.pop('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}list-level', None)
                style.attrib['{urn:oasis:names:tc:opendocument:xmlns:style:1.0}default-outline-level'] = level
            elif name == '通用多層清單':
                style.attrib.pop('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}list-style-name', None)
                style.attrib.pop('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}list-level', None)
                style.attrib['{urn:oasis:names:tc:opendocument:xmlns:style:1.0}default-outline-level'] = '5'

        working_outline_style = ET.fromstring(OUTLINE_STYLE_XML)
        outline_style = root.find('.//text:outline-style', ns)
        if outline_style is not None:
            parent_map = {c: p for p in root.iter() for c in p}
            parent_node = parent_map.get(outline_style)
            if parent_node is not None:
                idx = list(parent_node).index(outline_style)
                parent_node.remove(outline_style)
                parent_node.insert(idx, working_outline_style)
        else:
            styles_node = root.find('.//office:styles', ns)
            if styles_node is not None:
                styles_node.append(working_outline_style)

        for style in root.findall('.//style:style', ns):
            if style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family') == 'paragraph':
                props = style.find('.//style:paragraph-properties', ns)
                if props is None:
                    props = ET.SubElement(style, '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}paragraph-properties')
                props.attrib['{urn:oasis:names:tc:opendocument:xmlns:style:1.0}punctuation-wrap'] = 'simple'

        ln_config = root.find('.//text:linenumbering-configuration', ns)
        if ln_config is not None:
            ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}number-lines'] = 'true'
            ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}number-position'] = 'outside'
            ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:style:1.0}num-format'] = '1'
            ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}increment'] = '1'
        else:
            office_styles = root.find('.//office:styles', ns)
            if office_styles is not None:
                ln_config = ET.SubElement(office_styles, '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}linenumbering-configuration')
                ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}number-lines'] = 'true'
                ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}number-position'] = 'outside'
                ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:style:1.0}num-format'] = '1'
                ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}increment'] = '1'
                ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}offset'] = '1cm'
                ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}count-empty-lines'] = 'true'
                ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}count-in-floating-frames'] = 'false'
                ln_config.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}restart-on-page'] = 'true'

        return ET.tostring(root, encoding='utf-8')

def upgrade_odt(odt_path):
    odt_path = Path(odt_path)
    if not odt_path.exists():
        print(f"[ERROR] ODT file does not exist: {odt_path}")
        return False

    print(f"[START] Upgrading ODT: {odt_path.name}")
    temp_dir = tempfile.gettempdir()
    temp_output_path = Path(temp_dir) / (odt_path.name + ".upgrade.tmp")

    try:
        with zipfile.ZipFile(odt_path, 'r') as zin:
            content_xml = zin.read("content.xml")
            styles_xml = zin.read("styles.xml")
            other_files = {
                item.filename: zin.read(item.filename) 
                for item in zin.infolist() 
                if item.filename not in ("content.xml", "styles.xml")
            }
    except Exception as e:
        print(f"[ERROR] Failed to read ODT: {e}")
        return False

    print("  Processing XML structures...")
    try:
        new_content = process_content_xml(content_xml)
        new_styles = process_styles_xml(styles_xml)
    except Exception as e:
        print(f"[ERROR] Failed to process XML: {e}")
        return False

    print("  Rebuilding ODT package...")
    try:
        with zipfile.ZipFile(temp_output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            zout.writestr("content.xml", new_content)
            zout.writestr("styles.xml", new_styles)
            for fname, fdata in other_files.items():
                zout.writestr(fname, fdata)
        
        # 覆蓋原檔
        odt_path.unlink()
        temp_output_path.rename(odt_path)
        print(f"[OK] Successfully upgraded ODT: {odt_path.name}")
        return True
    except Exception as e:
        print(f"[ERROR] Failed to write ODT: {e}")
        if temp_output_path.exists():
            temp_output_path.unlink()
        return False

def main():
    parser = argparse.ArgumentParser(description="Upgrade existing ODT file(s) to support native Tab/Shift+Tab outline level demote/promote.")
    parser.add_argument("path", help="Path to the existing ODT file or directory containing ODT files.")
    args = parser.parse_args()
    
    target_path = Path(args.path)
    if not target_path.exists():
        print(f"[ERROR] Path does not exist: {target_path}")
        return

    if target_path.is_dir():
        odt_files = list(target_path.glob("*.odt"))
        if not odt_files:
            print(f"[INFO] No .odt files found in directory: {target_path}")
            return
        print(f"[INFO] Found {len(odt_files)} ODT files in directory: {target_path}")
        success_count = 0
        for f in odt_files:
            try:
                if upgrade_odt(f):
                    success_count += 1
            except Exception as e:
                print(f"[ERROR] Failed to upgrade {f.name}: {e}")
        print(f"[SUMMARY] Processed {len(odt_files)} files. Success: {success_count}, Failed: {len(odt_files) - success_count}")
    else:
        upgrade_odt(target_path)

if __name__ == "__main__":
    main()
