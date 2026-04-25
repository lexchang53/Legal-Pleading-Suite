import zipfile
import re
import os
from pathlib import Path

def fix_odt(odt_path):
    tmp_path = odt_path.with_suffix(".tmp.odt")
    
    target_parents = (
        '通用_5f_層級1', '通用_5f_層級2', '通用_5f_層級3', '通用_5f_層級4',
        '通用_層級1', '通用_層級2', '通用_層級3', '通用_層級4',
    )
    
    # Mapping style names to levels
    level_map = {
        '通用_5f_層級1': '0', '通用_5f_層級2': '1', '通用_5f_層級3': '2', '通用_5f_層級4': '3',
        '通用_層級1': '0', '通用_層級2': '1', '通用_層級3': '2', '通用_層級4': '3',
    }

    with zipfile.ZipFile(odt_path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            
            if item.filename in ["content.xml", "styles.xml"]:
                text = data.decode("utf-8", errors="replace")
                
                # 1. Fix Hanging Punctuation
                text = text.replace('style:punctuation-wrap="hanging"', 'style:punctuation-wrap="simple"')
                text = text.replace('style:hanging-punctuation="true"', 'style:hanging-punctuation="false"')
                
                # 2. Fix List Styles in content.xml
                if item.filename == "content.xml":
                    # Remove style:list-style-name from automatic styles that have target parents
                    def _strip_list_attr(m):
                        tag = m.group(0)
                        parent_match = re.search(r'style:parent-style-name="([^"]*)"', tag)
                        if not parent_match: return tag
                        if parent_match.group(1) in target_parents:
                            return re.sub(r'\s+style:list-style-name="[^"]*"', '', tag)
                        return tag

                    text = re.sub(r'<style:style\b[^>]*style:family="paragraph"[^>]*>', _strip_list_attr, text)
                
                # 3. Fix Line Numbering in styles.xml
                if item.filename == "styles.xml":
                    ln_config = (
                        '<text:linenumbering-configuration '
                        'text:number-lines="true" text:offset="1cm" style:num-format="1" '
                        'text:count-empty-lines="true" text:count-in-floating-frames="false" '
                        'text:restart-on-page="true" text:number-position="outside" text:increment="1"/>'
                    )
                    if 'text:number-position="outside"' not in text:
                        if '<text:linenumbering-configuration' in text:
                            text = re.sub(r'<text:linenumbering-configuration\b[^>]*?/?>', ln_config, text, count=1)
                        else:
                            for marker in ["</office:styles>", "</office:automatic-styles>"]:
                                if marker in text:
                                    text = text.replace(marker, ln_config + marker, 1)
                                    break
                
                data = text.encode("utf-8")
            
            zout.writestr(item, data)
            
    odt_path.unlink()
    tmp_path.rename(odt_path)
    print("Fixed ODT successfully (XML only)")

if __name__ == "__main__":
    import sys
    fix_odt(Path(sys.argv[1]))
