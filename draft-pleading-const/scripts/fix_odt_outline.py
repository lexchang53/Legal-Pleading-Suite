#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import zipfile
import re

OUTLINE_STYLE_XML = (
    '<text:outline-style style:name="Outline">'
    '<text:outline-level-style text:level="1" text:style-name="書狀_5f_層級1" '
    'loext:num-list-format="%1%、" style:num-suffix="、" style:num-format="一, 二, 三, ...">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="1.5cm" fo:text-indent="-1.5cm" fo:margin-left="1.5cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="2" text:style-name="書狀_5f_層級2" '
    'loext:num-list-format="(%2%)" style:num-suffix=" " style:num-format="一, 二, 三, ...">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="2.25cm" fo:text-indent="-1.5cm" fo:margin-left="2.25cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="3" text:style-name="書狀_5f_層級3" '
    'loext:num-list-format="%3%." style:num-suffix="." style:num-format="1, 2, 3, ...">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="3cm" fo:text-indent="-1cm" fo:margin-left="3cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '<text:outline-level-style text:level="4" text:style-name="書狀_5f_層級4" '
    'loext:num-list-format="(%4%)" style:num-suffix=" " style:num-format="1, 2, 3, ...">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" '
    'text:list-tab-stop-position="3.75cm" fo:text-indent="-1.25cm" fo:margin-left="3.75cm"/>'
    '</style:list-level-properties>'
    '</text:outline-level-style>'
    '</text:outline-style>'
)

LEVEL_STYLE_MAP = {
    "通用_5f_層級1": 1, "通用_5f_層級2": 2,
    "通用_5f_層級3": 3, "通用_5f_層級4": 4,
    "書狀_5f_層級1": 1, "書狀_5f_層級2": 2,
    "書狀_5f_層級3": 3, "書狀_5f_層級4": 4,
}

def enforce_outline_level(match):
    h_tag = match.group(0)
    s_m = re.search(r'text:style-name="([^"]+)"', h_tag)
    if s_m:
        sname = s_m.group(1)
        target_lvl = LEVEL_STYLE_MAP.get(sname)
        if target_lvl is not None:
            if 'text:outline-level=' in h_tag:
                h_tag = re.sub(r'text:outline-level="\d+"', f'text:outline-level="{target_lvl}"', h_tag)
            else:
                h_tag = re.sub(r'(<text:h\b[^>]*?)(\s*>)', rf'\1 text:outline-level="{target_lvl}"\2', h_tag)
            return h_tag
    return h_tag

def main():
    if len(sys.argv) < 2:
        print("Usage: python fix_odt_outline.py <input.odt>")
        sys.exit(1)
    
    odt_path = sys.argv[1]
    if not os.path.exists(odt_path):
        print(f"File not found: {odt_path}")
        sys.exit(1)
    
    temp_odt = odt_path + ".tmp"
    
    try:
        with zipfile.ZipFile(odt_path, 'r') as zin, zipfile.ZipFile(temp_odt, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'content.xml':
                    text = data.decode('utf-8')
                    text = re.sub(r'<text:h\b[^>]*>', enforce_outline_level, text)
                    data = text.encode('utf-8')
                elif item.filename == 'styles.xml':
                    text = data.decode('utf-8')
                    if re.search(r'<text:outline-style\b', text):
                        text = re.sub(
                            r'<text:outline-style\b[^>]*>.*?</text:outline-style>',
                            OUTLINE_STYLE_XML,
                            text,
                            flags=re.DOTALL
                        )
                    else:
                        text = text.replace(
                            '</office:styles>',
                            OUTLINE_STYLE_XML + '\n</office:styles>',
                            1
                        )
                    data = text.encode('utf-8')
                zout.writestr(item, data)
        # Replace original file with retry for Windows delayed locks
        for attempt in range(5):
            try:
                os.replace(temp_odt, odt_path)
                break
            except PermissionError:
                if attempt < 4:
                    import time
                    time.sleep(1)
                else:
                    raise
        print(f"[INFO] ODT Outline Levels fixed: {odt_path}")
    except Exception as e:
        if os.path.exists(temp_odt):
            os.remove(temp_odt)
        print(f"[ERROR] Failed to fix ODT: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
