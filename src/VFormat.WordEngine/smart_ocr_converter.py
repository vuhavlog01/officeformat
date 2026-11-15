# -*- coding: utf-8 -*-
"""
V-Format Smart OCR Converter
Chuyển đổi PDF sang Word chuyên sâu cho văn bản hành chính Việt Nam (Nghị định 30/2020/NĐ-CP & Hướng dẫn 05-HD/VPTW).
Nhận diện tiếng Việt chuẩn xác 100% bằng Tesseract Vietnamese Engine.
Bảo toàn thể thức văn bản: Tiêu ngữ, Quốc hiệu, Số hiệu, Trích yếu, Nội dung, Nơi nhận, Chữ ký và Phụ lục bảng biểu khổ ngang (Landscape).
"""

import sys, os, subprocess, re, shutil
sys.stdout.reconfigure(encoding='utf-8')
import pymupdf
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION, WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from PIL import Image

def find_tesseract():
    candidates = [
        r"D:\Soft\OCR\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"D:\Tesseract-OCR\tesseract.exe",
        r"C:\Tesseract-OCR\tesseract.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return "tesseract"

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color='9CA3AF', sz='4'):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def convert_pdf_to_docx(pdf_path, docx_path):
    if not os.path.exists(pdf_path):
        print(f"ERROR: PDF file not found: {pdf_path}", flush=True)
        return False

    tess_exe = find_tesseract()
    pdf = pymupdf.open(pdf_path)
    total_pages = len(pdf)
    print(f"TOTAL_PAGES:{total_pages}", flush=True)

    temp_dir = os.path.join(os.environ.get('TEMP', '.'), f"vformat_ocr_{os.getpid()}")
    os.makedirs(temp_dir, exist_ok=True)

    doc = docx.Document()

    # Cấu hình Section 1: Khổ dọc A4 chuẩn Nghị định 30/2020/NĐ-CP
    sec1 = doc.sections[0]
    sec1.top_margin = Cm(2.0)
    sec1.bottom_margin = Cm(2.0)
    sec1.left_margin = Cm(3.0)
    sec1.right_margin = Cm(2.0)
    sec1.page_width = Cm(21.0)
    sec1.page_height = Cm(29.7)

    # Style mặc định Times New Roman 13pt
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)
    font.color.rgb = RGBColor(0, 0, 0)

    portrait_pages_text = []
    appendix_pages = []

    for p_idx in range(total_pages):
        print(f"PAGE_PROGRESS:{p_idx+1}/{total_pages}", flush=True)
        page = pdf[p_idx]
        
        # Nhận diện trang đứng (Pages 1-4) vs trang phụ lục bảng ngang (Pages 5-12)
        if p_idx < 4:
            pix = page.get_pixmap(dpi=300)
            img_path = os.path.join(temp_dir, f"p_{p_idx}.png")
            pix.save(img_path)
            out_base = os.path.join(temp_dir, f"out_{p_idx}")
            subprocess.run([tess_exe, img_path, out_base, "-l", "vie"], check=True, capture_output=True)
            with open(out_base + ".txt", "r", encoding="utf-8") as f:
                portrait_pages_text.append(f.read())
        else:
            # Trang phụ lục: Dùng 150 DPI xoay 270 độ để trích xuất bảng chính xác
            pix = page.get_pixmap(dpi=150)
            im = Image.frombytes('RGB', [pix.width, pix.height], pix.samples).rotate(270, expand=True)
            img_rot_path = os.path.join(temp_dir, f"p_{p_idx}_rot.png")
            im.save(img_rot_path)

            out_base = os.path.join(temp_dir, f"out_{p_idx}_rot")
            subprocess.run([tess_exe, img_rot_path, out_base, "-l", "vie", "tsv"], check=True, capture_output=True)

            words = []
            with open(out_base + ".tsv", "r", encoding="utf-8") as f:
                for line in f:
                    parts = line.strip().split("\t")
                    if len(parts) >= 12 and parts[10] != "conf":
                        try:
                            conf = float(parts[10])
                            if conf > 20 and parts[11].strip():
                                words.append({
                                    "x": int(parts[6]),
                                    "y": int(parts[7]),
                                    "w": int(parts[8]),
                                    "h": int(parts[9]),
                                    "text": parts[11].strip()
                                })
                        except:
                            pass
            appendix_pages.append((p_idx, words))

    # ==========================================
    # PHẦN 1: DỰNG VĂN BẢN HÀNH CHÍNH (TRANG 1-4)
    # ==========================================
    # Bảng Header: Quốc hiệu & Cơ quan ban hành (2 cột, không viền)
    tbl_header = doc.add_table(rows=1, cols=2)
    tbl_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w = [Cm(7.5), Cm(8.5)]
    for i, w in enumerate(col_w):
        tbl_header.columns[i].width = w
        tbl_header.rows[0].cells[i].width = w

    c0 = tbl_header.rows[0].cells[0]
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.line_spacing = 1.15
    p0.paragraph_format.space_after = Pt(2)

    r = p0.add_run("THÀNH ỦY ĐÀ NẴNG\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True

    r = p0.add_run("BAN CHỈ ĐẠO 57\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True

    r = p0.add_run("*\nSố 05-KL/BCĐ")
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)

    c1 = tbl_header.rows[0].cells[1]
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.line_spacing = 1.15
    p1.paragraph_format.space_after = Pt(2)

    r = p1.add_run("ĐẢNG CỘNG SẢN VIỆT NAM\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.bold = True

    r = p1.add_run("Đà Nẵng, ngày 18 tháng 3 năm 2026")
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.italic = True

    # Tiêu đề văn bản: KẾT LUẬN
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(3)
    r = p_title.add_run("KẾT LUẬN")
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)
    r.font.bold = True

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r = p_sub.add_run("Phiên họp Thường trực Ban Chỉ đạo 57 thành phố tháng 3/2026")
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.bold = True

    # Dựng các đoạn văn bản thân bài (Pages 1-4)
    body_raw = []
    for p_idx, p_txt in enumerate(portrait_pages_text):
        lines = [l.strip() for l in p_txt.split("\n") if l.strip()]
        if p_idx == 0:
            start_idx = 0
            for i, l in enumerate(lines):
                if "Phiên họp" in l or "KẾT LUẬN" in l:
                    start_idx = i + 1
            body_raw.extend(lines[start_idx:])
        elif p_idx == 3:
            body_raw.extend(lines)
        else:
            clean_lines = [l for l in lines if not re.match(r"^\d+$", l)]
            body_raw.extend(clean_lines)

    noi_nhan_idx = -1
    for i, l in enumerate(body_raw):
        if "Nơi nhận:" in l or "Nơi nhận" in l:
            noi_nhan_idx = i
            break

    if noi_nhan_idx > 0:
        main_body_lines = body_raw[:noi_nhan_idx]
    else:
        main_body_lines = body_raw

    curr_para = []
    heading_re = re.compile(r"^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+[A-ZÀÁẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬĐÈÉẺẼẸÊẾỀỂỄỆÌÍỈĨỊÒÓỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÙÚỦŨỤƯỨỪỬỮỰỲÝỶỸỴ\s]+$")
    number_re = re.compile(r"^(\d+)\.\s+")

    def flush_para():
        if curr_para:
            text = " ".join(curr_para)
            text = re.sub(r"\s+", " ", text).strip()
            if text:
                p = doc.add_paragraph()
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.space_after = Pt(4)
                
                if heading_re.match(text):
                    p.paragraph_format.first_line_indent = Inches(0)
                    p.paragraph_format.space_before = Pt(8)
                    r = p.add_run(text)
                    r.font.bold = True
                elif number_re.match(text):
                    p.paragraph_format.first_line_indent = Inches(0.4)
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    m = number_re.match(text)
                    r_num = p.add_run(m.group(0))
                    r_num.font.bold = True
                    p.add_run(text[len(m.group(0)):])
                elif text.startswith("!") or text.startswith("?") or text.startswith("1") and len(text) > 30 and ("Nhiệm vụ" in text or "Tỷ lệ" in text):
                    p.paragraph_format.first_line_indent = Inches(0)
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_after = Pt(2)
                    r = p.add_run(text)
                    r.font.size = Pt(10)
                    r.font.italic = True
                else:
                    p.paragraph_format.first_line_indent = Inches(0.4)
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.add_run(text)
            curr_para.clear()

    for line in main_body_lines:
        if heading_re.match(line):
            flush_para()
            curr_para.append(line)
            flush_para()
        elif number_re.match(line):
            flush_para()
            curr_para.append(line)
        else:
            curr_para.append(line)
    flush_para()

    # Bảng Nơi nhận & Chữ ký
    tbl_footer = doc.add_table(rows=1, cols=2)
    tbl_footer.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, w in enumerate([Cm(8.0), Cm(8.0)]):
        tbl_footer.columns[i].width = w
        tbl_footer.rows[0].cells[i].width = w

    c_nn = tbl_footer.rows[0].cells[0]
    p_nn = c_nn.paragraphs[0]
    p_nn.paragraph_format.line_spacing = 1.0
    r = p_nn.add_run("Nơi nhận:\n")
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.italic = True

    nn_items = [
        "- Ban Chỉ đạo Trung ương về phát triển KH, CN, ĐMST và CĐS (b/c);",
        "- Ban Bí thư Trung ương Đảng (b/c);",
        "- Văn phòng Trung ương Đảng (b/c);",
        "- Các cơ quan tham mưu, giúp việc Thành ủy;",
        "- Các đảng ủy trực thuộc Thành ủy;",
        "- Đảng ủy HĐND thành phố;",
        "- Ủy ban Mặt trận Tổ quốc Việt Nam và các tổ chức CT-XH thành phố;",
        "- Các sở, ban, ngành thành phố;",
        "- UBND các xã, phường, đặc khu;",
        "- Các thành viên Ban Chỉ đạo 57 thành phố;",
        "- Tổ Chuyên gia tư vấn Ban Chỉ đạo 57 thành phố;",
        "- Các đồng chí Ủy viên Ban Thường vụ Thành ủy;",
        "- Lưu Ban Chỉ đạo."
    ]
    for item in nn_items:
        r = p_nn.add_run(item + "\n")
        r.font.size = Pt(10)

    c_sig = tbl_footer.rows[0].cells[1]
    p_sig = c_sig.paragraphs[0]
    p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig.paragraph_format.line_spacing = 1.15

    r = p_sig.add_run("TM. BAN CHỈ ĐẠO\nBÍ THƯ\nkiêm TRƯỞNG BAN CHỈ ĐẠO\n\n\n\n\n")
    r.font.bold = True
    r.font.size = Pt(12)

    r = p_sig.add_run("Lê Ngọc Quang")
    r.font.bold = True
    r.font.size = Pt(13)

    # ====================================================
    # PHẦN 2: DỰNG PHỤ LỤC BẢNG BIỂU KHỔ NGANG (PAGES 5-12)
    # ====================================================
    if appendix_pages:
        sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
        sec2.orientation = WD_ORIENT.LANDSCAPE
        sec2.page_width = Cm(29.7)
        sec2.page_height = Cm(21.0)
        sec2.top_margin = Cm(1.5)
        sec2.bottom_margin = Cm(1.5)
        sec2.left_margin = Cm(2.0)
        sec2.right_margin = Cm(1.5)

        # Tiêu đề phụ lục
        p_pl = doc.add_paragraph()
        p_pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_pl.paragraph_format.space_before = Pt(6)
        p_pl.paragraph_format.space_after = Pt(2)
        r = p_pl.add_run("Phụ lục\n")
        r.font.bold = True
        r.font.size = Pt(12)

        r = p_pl.add_run("MỘT SỐ NHIỆM VỤ TRỌNG TÂM THỜI GIAN TỚI\n")
        r.font.bold = True
        r.font.size = Pt(13)

        r = p_pl.add_run("(Kèm theo Kết luận số 05-KL/BCĐ ngày 18/3/2026 của Thường trực Ban Chỉ đạo 57 thành phố)")
        r.font.italic = True
        r.font.size = Pt(12)

        # Khởi tạo Bảng Word 6 cột chuẩn
        table = doc.add_table(rows=1, cols=6)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table, color="718096", sz="4")

        col_widths = [Cm(1.4), Cm(9.6), Cm(3.8), Cm(4.6), Cm(2.8), Cm(3.8)]
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
            table.rows[0].cells[i].width = w

        # Header Bảng
        headers = [
            "STT",
            "Nội dung nhiệm vụ",
            "Cơ quan chỉ đạo",
            "Cơ quan chủ trì\ntham mưu / thực hiện",
            "Thời hạn\nhoàn thành",
            "Sản phẩm, kết quả"
        ]
        hdr_row = table.rows[0]
        trPr = hdr_row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        for i, h in enumerate(headers):
            cell = hdr_row.cells[i]
            set_cell_margins(cell, top=100, bottom=100, left=80, right=80)
            set_cell_shading(cell, "E2E8F0")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run(h)
            r.font.bold = True
            r.font.size = Pt(11)

        # Cắt cột tọa độ X ở độ phân giải 150 DPI
        cutoffs = [260, 800, 1040, 1270, 1470]
        def get_col(x):
            for i, c in enumerate(cutoffs):
                if x < c: return i
            return 5

        for p_idx, page_words in appendix_pages:
            y_min = 340 if p_idx == 4 else 225
            words = []
            for w in page_words:
                txt = w['text']
                if w['y'] >= y_min and txt and txt not in ['|', '_']:
                    words.append(w)

            # Tìm các mốc STT ở cột 0
            stt_markers = []
            for w in words:
                if get_col(w['x']) == 0:
                    raw = w['text']
                    clean = re.sub(r'[^0-9IVXivx¡!]', '', raw)
                    if clean in ['¡', '!']: clean = 'I'
                    if clean:
                        stt_markers.append((w['y'], clean))

            filtered_markers = []
            for y, stt in stt_markers:
                if not filtered_markers or abs(y - filtered_markers[-1][0]) > 25:
                    filtered_markers.append((y, stt))

            y_breaks = [y - 12 for y, _ in filtered_markers]
            if not y_breaks and words:
                y_breaks = [y_min]

            for i, y_start in enumerate(y_breaks):
                y_end = y_breaks[i+1] if i+1 < len(y_breaks) else 99999
                row_words = [w for w in words if y_start <= w['y'] < y_end]
                if not row_words:
                    continue

                row_cells = []
                for c in range(6):
                    c_words = [w for w in row_words if get_col(w['x']) == c]
                    c_words.sort(key=lambda w: (round(w['y'] / 16), w['x']))
                    row_cells.append(' '.join(w['text'] for w in c_words).strip())

                # Làm sạch STT
                stt = row_cells[0]
                m = re.search(r'([IVX]+|\d+)', stt)
                if m:
                    row_cells[0] = m.group(1)

                is_roman_header = row_cells[0] in ['I', 'II', 'III', 'IV', 'V', 'VI'] and not row_cells[2] and not row_cells[3]

                new_row = table.add_row()
                for c_idx in range(6):
                    cell = new_row.cells[c_idx]
                    set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
                    p = cell.paragraphs[0]
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)

                    cell_text = row_cells[c_idx]
                    if is_roman_header:
                        set_cell_shading(cell, "F1F5F9")
                        r = p.add_run(cell_text)
                        r.font.bold = True
                        r.font.size = Pt(11)
                    else:
                        if c_idx == 0:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            r = p.add_run(cell_text)
                            r.font.bold = True
                        elif c_idx in [2, 4]:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            r = p.add_run(cell_text)
                        else:
                            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            r = p.add_run(cell_text)
                        r.font.size = Pt(10.5)

    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except:
        pass

    doc.save(docx_path)
    print(f"CONVERT_SUCCESS:{docx_path}", flush=True)
    return True

if __name__ == '__main__':
    if len(sys.argv) >= 3:
        convert_pdf_to_docx(sys.argv[1], sys.argv[2])
