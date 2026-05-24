import argparse
import os
import sys
import fitz  # PyMuPDF
import re
from collections import Counter

def is_page_number_like(text):
    """
    判定字串是否極有可能是頁碼，如果是，則回傳 True。
    """
    text = text.strip()
    
    # 1. 排除純數字 (如 "1", "23")
    if text.isdigit():
        return True
        
    # 2. 排除極短且只包含數字、空格、斜線或減號的字串 (如 "1 / 5", " - 2 -", "03")
    if len(text) <= 5 and re.match(r'^[\d\s\-/]+$', text):
        return True
        
    # 3. 排除常見的頁碼文字格式
    page_patterns = [
        r'^第\s*\d+\s*頁$',                         # "第 1 頁"
        r'^第\s*\d+\s*頁\s*/\s*共\s*\d+\s*頁$',     # "第 1 頁 / 共 10 頁"
        r'^[Pp]age\s*\d+$',                         # "Page 1"
        r'^[Pp]age\s*\d+\s*of\s*\d+$'               # "page 1 of 10"
    ]
    for pattern in page_patterns:
        if re.match(pattern, text, re.IGNORECASE):
            return True
            
    return False

def remove_watermarks(input_path, output_path, margin_top=45.0, margin_bottom=45.0, freq_threshold=3):
    """
    動態適應橫/直式頁面，移除角落姓名與日期文字，並無損清除角落小圖片（如天平浮水印）的引用與指令。
    """
    if not os.path.exists(input_path):
        print(f"錯誤：輸入檔案不存在：{input_path}", file=sys.stderr)
        return False

    try:
        doc = fitz.open(input_path)
        total_pages = len(doc)
        print(f"正在載入 PDF: {input_path} (共 {total_pages} 頁)")
        
        # ==========================================
        # 第一階段：掃描邊緣文字並統計出現頻率，同時偵測浮水印小圖片名稱
        # ==========================================
        print("步驟 1/3: 正在掃描分析邊緣文字與圖片浮水印...")
        edge_text_counter = Counter()
        watermark_images_by_page = {}  # page_num -> list of image names to delete
        has_judicial_stamp = False
        
        # 司法院專屬浮水印白名單（直接刪除，不需統計頻率）
        judicial_keywords = ["司法院線上閱卷系統作業平台", "線上閱卷系統"]
        
        for page_num in range(total_pages):
            page = doc[page_num]
            page_height = page.rect.height
            
            # 動態計算此頁的頂部與底部邊緣閾值
            top_threshold = margin_top
            bottom_threshold = page_height - margin_bottom
            
            # 1. 偵測文字浮水印
            blocks = page.get_text("blocks")
            for b in blocks:
                x0, y0, x1, y1, text, _, _ = b
                cleaned_text = text.strip()
                if not cleaned_text:
                    continue
                
                # 判定文字塊是否在邊緣
                is_top_edge = y1 < top_threshold
                is_bottom_edge = y0 > bottom_threshold
                
                if is_top_edge or is_bottom_edge:
                    if any(kw in cleaned_text for kw in judicial_keywords):
                        has_judicial_stamp = True
                        continue
                    if is_page_number_like(cleaned_text) or len(cleaned_text) <= 1:
                        continue
                    edge_text_counter[cleaned_text] += 1
            
            # 2. 偵測圖片浮水印 (例如四個角落的天平小圖片，通常小於 100x100 pt)
            images = page.get_images(full=True)
            for img_info in images:
                width = img_info[2]
                height = img_info[3]
                img_name = img_info[7]
                
                # 篩選尺寸在 10x10 到 100x100 之間的小型浮水印圖片
                if 10 <= width <= 100 and 10 <= height <= 100:
                    if page_num not in watermark_images_by_page:
                        watermark_images_by_page[page_num] = []
                    if img_name not in watermark_images_by_page[page_num]:
                        watermark_images_by_page[page_num].append(img_name)
        
        # 篩選出在角落重複出現大於等於閾值的姓名浮水印
        watermark_signatures = {
            text for text, count in edge_text_counter.items() 
            if count >= freq_threshold
        }
        
        # 列印偵測結果
        print("偵測到的浮水印特徵：")
        if watermark_signatures:
            for sig in watermark_signatures:
                print(f"  - 角落文字: \"{sig}\" (出現 {edge_text_counter[sig]} 次)")
        if has_judicial_stamp:
            print("  - 底部防偽: 司法院線上閱卷系統作業平台 (動態時間戳記)")
        
        total_img_watermarks = sum(len(names) for names in watermark_images_by_page.values())
        if total_img_watermarks > 0:
            print(f"  - 角落圖片: 偵測到 {total_img_watermarks} 處小型圖片浮水印（天平等圖示）")
            
        # ==========================================
        # 第二階段：移除圖片浮水印的資源引用與內容流指令（無損大圖）
        # ==========================================
        if watermark_images_by_page:
            print("步驟 2/3: 正在清除圖片浮水印之資源與繪製指令（無損背景）...")
            img_removed_count = 0
            for page_num, img_names in watermark_images_by_page.items():
                page = doc[page_num]
                res_val = doc.xref_get_key(page.xref, "Resources")
                res_xref = None
                if isinstance(res_val, tuple) and res_val[0] == 'xref':
                    res_xref = int(res_val[1].split()[0])
                elif isinstance(res_val, str) and ' 0 R' in res_val:
                    res_xref = int(res_val.split()[0])
                    
                if res_xref:
                    for name in img_names:
                        # 1. 從頁面資源字典 XObject 中解綁此圖片
                        doc.xref_set_key(res_xref, f"XObject/{name}", "null")
                        
                        # 2. 從該頁的內容流中清除它的繪製指令 /name Do
                        contents = page.get_contents()
                        for c_xref in contents:
                            stream_data = doc.xref_stream(c_xref)
                            target_op = f"/{name} Do".encode('latin1')
                            if target_op in stream_data:
                                new_stream = stream_data.replace(target_op, b"")
                                doc.update_stream(c_xref, new_stream)
                        img_removed_count += 1
            print(f"圖片浮水印清除完畢，共處理 {img_removed_count} 個圖片資源。")

        # ==========================================
        # 第三階段：擦除文字浮水印
        # ==========================================
        print("步驟 3/3: 正在執行精準文字浮水印擦除...")
        text_removed_count = 0
        
        for page_num in range(total_pages):
            page = doc[page_num]
            page_height = page.rect.height
            top_threshold = margin_top
            bottom_threshold = page_height - margin_bottom
            
            blocks = page.get_text("blocks")
            page_has_redaction = False
            
            for b in blocks:
                x0, y0, x1, y1, text, _, _ = b
                cleaned_text = text.strip()
                if not cleaned_text:
                    continue
                
                is_top_edge = y1 < top_threshold
                is_bottom_edge = y0 > bottom_threshold
                
                if is_top_edge or is_bottom_edge:
                    # 條件 1: 包含司法院防偽關鍵字，直接擦除
                    is_judicial_stamp = any(kw in cleaned_text for kw in judicial_keywords)
                    
                    # 條件 2: 屬於重複出現的角落姓名
                    is_name_watermark = (cleaned_text in watermark_signatures)
                    
                    if is_judicial_stamp or is_name_watermark:
                        rect = fitz.Rect(x0, y0, x1, y1)
                        page.add_redact_annot(rect, fill=None)
                        page_has_redaction = True
                        text_removed_count += 1
            
            if page_has_redaction:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
                
            if (page_num + 1) % 50 == 0 or (page_num + 1) == total_pages:
                print(f"已處理進度: {page_num + 1}/{total_pages} 頁...")

        # 儲存結果並最佳化
        print("正在最佳化並另存 PDF 檔案...")
        doc.save(
            output_path, 
            garbage=4,
            deflate=True,
            clean=True
        )
        doc.close()
        print(f"處理完成！共移除了 {text_removed_count} 處文字與 {total_img_watermarks} 處圖片浮水印。")
        return True

    except Exception as e:
        print(f"處理 PDF 時發生未預期的錯誤: {str(e)}", file=sys.stderr)
        return False

def main():
    parser = argparse.ArgumentParser(description="台灣法院電子卷證 PDF 邊緣浮水印物理去浮水印工具")
    parser.add_argument("-i", "--input", required=True, help="輸入的 PDF 檔案路徑")
    parser.add_argument("-o", "--output", required=True, help="輸出的 PDF 檔案路徑")
    parser.add_argument("-t", "--top", type=float, default=45.0, help="頂部邊緣判定距離 (pt)，預設 45.0")
    parser.add_argument("-b", "--bottom", type=float, default=45.0, help="底部邊緣判定距離 (pt)，預設 45.0")
    parser.add_argument("-f", "--freq", type=int, default=3, help="角落姓名重複次數閾值，預設 3")
    
    args = parser.parse_args()
    
    success = remove_watermarks(args.input, args.output, args.top, args.bottom, args.freq)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
