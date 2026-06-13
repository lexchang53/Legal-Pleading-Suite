#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import os
import re
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


def register_all_namespaces(xml_string):
    namespaces = re.findall(r'xmlns:([^=]+)="([^"]+)"', xml_string)
    for prefix, uri in namespaces:
        ET.register_namespace(prefix, uri)
    # 手動註冊常用命名空間以確保輸出標籤乾淨
    ET.register_namespace('style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0')
    ET.register_namespace('fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0')
    ET.register_namespace('text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0')
    ET.register_namespace('office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0')


def fix_odt(odt_path):
    if not os.path.exists(odt_path):
        raise FileNotFoundError(f"ODT file not found: {odt_path}")
        
    print(f"Post-processing ODT (Constitutional Pleading Restart Numbering): {odt_path}...")
    
    # 1. 解開 ODT 檔案進行 XML 精準修正 (僅修改 content.xml 中的編號重設)
    temp_dir = tempfile.gettempdir()
    temp_output_path = os.path.join(temp_dir, os.path.basename(odt_path) + ".tmp")
    
    try:
        with zipfile.ZipFile(odt_path, 'r') as zin:
            content_xml = zin.read("content.xml").decode("utf-8")
            other_files = {
                item.filename: zin.read(item.filename) 
                for item in zin.infolist() 
                if item.filename != "content.xml"
            }
    except Exception as e:
        print(f"  Error reading ODT: {e}")
        return False

    ns = {
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
    }

    # --- 修改 content.xml (區塊後一級論點編號重設) ---
    register_all_namespaces(content_xml)
    content_root = ET.fromstring(content_xml)

    # 1. 識別一級論點段落與區塊標題
    level1_styles = {'通用_5f_層級1', '通用_層級1'}
    for style in content_root.findall('.//style:style', ns):
        name = style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name')
        parent = style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name')
        if parent in level1_styles:
            if name:
                level1_styles.add(name)
                
    section_styles = {'書狀_5f_區塊標題', '書狀_區塊標題'}
    for style in content_root.findall('.//style:style', ns):
        name = style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name')
        parent = style.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name')
        if parent in section_styles:
            if name:
                section_styles.add(name)

    restart_next = False
    cjk_numeral_pattern = re.compile(r'^[　\s]*[壹貳參肆伍陸柒捌玖拾]+[、]')
    section_keywords = ["聲請審查客體", "應受判決事項之聲明", "主要爭點", "聲請理由"]
    
    fixed_resets = 0
    for elem in content_root.iter():
        tag = elem.tag.split('}')[-1]
        if tag in ('h', 'p'):
            text_content = "".join(elem.itertext()).strip()
            style_name = elem.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name')
            
            is_section = False
            if style_name in section_styles:
                is_section = True
            elif any(kw in text_content for kw in section_keywords):
                is_section = True
            elif cjk_numeral_pattern.match(text_content):
                is_section = True
                
            if is_section:
                restart_next = True
                print(f"  [content.xml] Detected heading section: '{text_content[:20]}' -> next L1 will restart")
            elif style_name in level1_styles:
                if restart_next:
                    elem.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}restart-numbering'] = 'true'
                    elem.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}start-value'] = '1'
                    print(f"  [content.xml] Restarted numbering on level-1: '{text_content[:20]}'")
                    fixed_resets += 1
                    restart_next = False

    new_content_xml = ET.tostring(content_root, encoding='utf-8').decode('utf-8')
    print(f"  [content.xml] Applied {fixed_resets} numbering restarts.")

    # --- 2. 寫入臨時 ODT，然後覆蓋原檔案 ---
    try:
        with zipfile.ZipFile(temp_output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            zout.writestr("content.xml", new_content_xml.encode("utf-8"))
            for fname, fdata in other_files.items():
                zout.writestr(fname, fdata)
        
        if os.path.exists(odt_path):
            os.remove(odt_path)
        os.rename(temp_output_path, odt_path)
        print("  Post-processing completed successfully.")
        return True
    except Exception as e:
        print(f"  Error writing output ODT: {e}")
        if os.path.exists(temp_output_path):
            os.remove(temp_output_path)
        return False


def main():
    parser = argparse.ArgumentParser(description="Fix ODT outline alignments and numbering resets.")
    parser.add_argument("odt", help="Path to ODT file to fix.")
    args = parser.parse_args()
    
    fix_odt(args.odt)


if __name__ == "__main__":
    main()
