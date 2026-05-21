#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_odt_tab.py
一鍵修復舊版 ODT 書狀的 Tab 升降級與數字編號。
支援修復單一 ODT 檔案或遞迴掃描整個資料夾。
"""

import os
import sys
import argparse
import zipfile
import shutil
from pathlib import Path
from lxml import etree

def qn(prefix, localname):
    nsmap = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    }
    return f"{{{nsmap[prefix]}}}{localname}"

LEVEL_MAP = {
    '通用_層級1': '1', '通用_層級2': '2', '通用_層級3': '3', '通用_層級4': '4',
    '通用_5f_層級1': '1', '通用_5f_層級2': '2', '通用_5f_層級3': '3', '通用_5f_層級4': '4',
}

OUTLINE_STYLE_XML = """<text:outline-style xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" style:name="Outline">
  <text:outline-level-style text:level="1" loext:num-list-format="%1%、" style:num-suffix="、" style:num-format="一, 二, 三, ...">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1cm" fo:text-indent="-1cm" fo:margin-left="1cm"/>
    </style:list-level-properties>
  </text:outline-level-style>
  <text:outline-level-style text:level="2" loext:num-list-format="(%2%)" style:num-prefix="(" style:num-suffix=")" style:num-format="一, 二, 三, ...">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.499cm" fo:text-indent="-1cm" fo:margin-left="1.499cm"/>
    </style:list-level-properties>
  </text:outline-level-style>
  <text:outline-level-style text:level="3" loext:num-list-format="%3%." style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.499cm" fo:text-indent="-0.499cm" fo:margin-left="1.499cm"/>
    </style:list-level-properties>
  </text:outline-level-style>
  <text:outline-level-style text:level="4" loext:num-list-format="(%4%)" style:num-prefix="(" style:num-suffix=")" style:num-format="1">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.998cm" fo:text-indent="-0.499cm" fo:margin-left="1.998cm"/>
    </style:list-level-properties>
  </text:outline-level-style>
  <text:outline-level-style text:level="5" loext:num-list-format="%5%" style:num-format="">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="nothing"/>
    </style:list-level-properties>
  </text:outline-level-style>
  <text:outline-level-style text:level="6" loext:num-list-format="%6%" style:num-format="">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="nothing"/>
    </style:list-level-properties>
  </text:outline-level-style>
  <text:outline-level-style text:level="7" loext:num-list-format="%7%" style:num-format="">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="nothing"/>
    </style:list-level-properties>
  </text:outline-level-style>
  <text:outline-level-style text:level="8" loext:num-list-format="%8%" style:num-format="">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="nothing"/>
    </style:list-level-properties>
  </text:outline-level-style>
  <text:outline-level-style text:level="9" loext:num-list-format="%9%" style:num-format="">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
      <style:list-level-label-alignment text:label-followed-by="nothing"/>
    </style:list-level-properties>
  </text:outline-level-style>
</text:outline-style>"""

def fix_single_odt(filepath: Path) -> bool:
    try:
        print(f"正在修復 ODT：{filepath} ...")
        tmp_path = filepath.with_suffix(f".fix_tmp_{os.getpid()}")
        
        # 解析內嵌的標準大綱樣式 XML
        working_outline_style = etree.fromstring(OUTLINE_STYLE_XML)

        with zipfile.ZipFile(filepath, 'r') as zin:
            with zipfile.ZipFile(tmp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'content.xml':
                        root = etree.fromstring(zin.read(item.filename))
                        auto_styles = {}
                        
                        # 建立自動樣式對應表並剝除 list-style-name 與 list-level 屬性
                        for style in root.findall('.//style:style', root.nsmap):
                            name = style.get(qn('style', 'name'))
                            parent = style.get(qn('style', 'parent-style-name'))
                            if name and parent:
                                if parent in LEVEL_MAP:
                                    auto_styles[name] = parent
                                if parent in LEVEL_MAP or parent == '通用多層清單':
                                    for attr in (qn('style', 'list-style-name'), qn('style', 'list-level')):
                                        if attr in style.attrib:
                                            del style.attrib[attr]

                        # 1. 剝除所有清單外衣 <text:list>
                        while True:
                            lists = root.findall('.//text:list', root.nsmap)
                            if not lists:
                                break
                            for lst in lists:
                                parent = lst.getparent()
                                if parent is None:
                                    continue
                                index = parent.index(lst)
                                for list_item in lst.findall('text:list-item', root.nsmap):
                                    for child in list(list_item):
                                        parent.insert(index, child)
                                        index += 1
                                parent.remove(lst)

                        # 2. 將通用層級轉換為 <text:h> 大綱標題
                        for p in root.findall('.//text:p', root.nsmap):
                            sname = p.get(qn('text', 'style-name'))
                            target_level = None
                            if sname in LEVEL_MAP:
                                target_level = LEVEL_MAP[sname]
                            elif sname in auto_styles and auto_styles[sname] in LEVEL_MAP:
                                target_level = LEVEL_MAP[auto_styles[sname]]
                                
                            if target_level:
                                p.tag = qn('text', 'h')
                                p.set(qn('text', 'outline-level'), target_level)
                                
                        # 3. 清理書籤括號
                        for bookmark in root.findall('.//text:bookmark-start', root.nsmap):
                            bookmark.getparent().remove(bookmark)
                        for bookmark in root.findall('.//text:bookmark-end', root.nsmap):
                            bookmark.getparent().remove(bookmark)
                        for bookmark in root.findall('.//text:bookmark', root.nsmap):
                            bookmark.getparent().remove(bookmark)

                        zout.writestr(item, etree.tostring(root, xml_declaration=True, encoding='UTF-8'))

                    elif item.filename == 'styles.xml':
                        root = etree.fromstring(zin.read(item.filename))
                        
                        # 4. 修正樣式的 outline-level 屬性，並移除清單關聯（包括通用多層清單父樣式）
                        for style in root.findall('.//style:style', root.nsmap):
                            name = style.get(qn('style', 'name'))
                            if name in LEVEL_MAP:
                                level = LEVEL_MAP[name]
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

                        # 5. 替換 outline-style 為台灣法律大綱格式
                        outline_style = root.find('.//text:outline-style', root.nsmap)
                        if outline_style is not None:
                            parent = outline_style.getparent()
                            index = parent.index(outline_style)
                            parent.remove(outline_style)
                            parent.insert(index, working_outline_style)
                        else:
                            styles_node = root.find('.//office:styles', root.nsmap)
                            if styles_node is not None:
                                styles_node.append(working_outline_style)
                                    
                        zout.writestr(item, etree.tostring(root, xml_declaration=True, encoding='UTF-8'))
                    else:
                        zout.writestr(item, zin.read(item.filename))
                        
        # 用修復後的檔案取代原檔案
        shutil.move(str(tmp_path), str(filepath))
        print("[OK] 修復成功！")
        return True
    except Exception as e:
        print(f"[FAIL] 修復失敗 ({filepath.name})：{e}")
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass
        return False

def main():
    parser = argparse.ArgumentParser(description="一鍵修復舊版 ODT 書狀的 Tab 升降級與數字編號")
    parser.add_argument("path", help="要修復的 ODT 檔案路徑，或包含 ODT 檔案的資料夾路徑")
    args = parser.parse_args()

    target_path = Path(args.path).resolve()
    if not target_path.exists():
        print(f"[FAIL] 找不到路徑：{target_path}")
        sys.exit(1)

    if target_path.is_file():
        if target_path.suffix.lower() != '.odt':
            print("[FAIL] 目標檔案不是 .odt 檔案！")
            sys.exit(1)
        success = fix_single_odt(target_path)
        sys.exit(0 if success else 1)
        
    elif target_path.is_dir():
        print(f"正在掃描資料夾：{target_path} ...")
        odt_files = list(target_path.rglob("*.odt"))
        if not odt_files:
            print("沒有找到任何 .odt 檔案。")
            sys.exit(0)
            
        print(f"共找到 {len(odt_files)} 個 ODT 檔案。")
        print("-" * 50)
        
        success_count = 0
        for f in odt_files:
            # 排除暫存檔
            if f.name.startswith("~$") or ".fix_tmp_" in f.name:
                continue
            if fix_single_odt(f):
                success_count += 1
                
        print("-" * 50)
        print(f"[完成] 批次處理完畢！成功修復: {success_count} / {len(odt_files)} 個檔案。")

if __name__ == '__main__':
    main()
