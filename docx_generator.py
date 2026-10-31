import os
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

FONT_NAME = "Times New Roman"

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Đặt lề trong của ô (twips: 20 twips = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    """
    Thiết lập viền tùy biến cho ô trong bảng (top, bottom, left, right).
    Ví dụ: set_cell_border(cell, bottom={'val': 'dotted', 'sz': '4'})
    """
    tcPr = cell._tc.get_or_add_tcPr()
    xml_parts = []
    for edge, conf in kwargs.items():
        val = conf.get('val', 'single')
        sz = conf.get('sz', '4')
        color = conf.get('color', 'auto')
        xml_parts.append(f'<w:{edge} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>')
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>{"".join(xml_parts)}</w:tcBorders>')
    tcPr.append(tcBorders)

def set_table_borders(table, color="000000", sz="4", val="single"):
    """Đặt viền chuẩn theo mẫu: viền ngoài nét đôi thanh đậm, viền trong dọc kép."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="thinThickLargeGap" w:sz="24" w:space="0" w:color="{color}"/>\n'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color}"/>\n'
        f'  <w:left w:val="thinThickLargeGap" w:sz="24" w:space="0" w:color="{color}"/>\n'
        f'  <w:right w:val="thinThickLargeGap" w:sz="24" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideV w:val="double" w:sz="4" w:space="0" w:color="{color}"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def make_row_header(row):
    """Lặp lại hàng tiêu đề bảng ở mỗi trang."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

def make_row_cant_split(row):
    """Không để hàng bị ngắt ngang trang."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def generate_weekly_schedule_docx(data, output_path, signature_path=None):
    """
    Tạo file Word Lịch công tác tuần đúng chuẩn thể thức văn bản Đảng
    và bám sát 100% mẫu thông báo thực tế.
    """
    doc = docx.Document()
    
    # 1. Cấu hình lề trang chuẩn A4
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.0)
    
    # Cấu hình Header trang (số trang ở góc trên hoặc giữa từ trang 2)
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Có thể để trống hoặc thêm page number nếu cần
    
    # 2. Tiêu ngữ Đảng (2 cột không viền)
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    
    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)
    cell_left.width = Cm(8.5)
    cell_right.width = Cm(9.9)
    
    # Cột trái
    p_cq_tren = cell_left.paragraphs[0]
    p_cq_tren.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cq_tren.paragraph_format.space_before = Pt(0)
    p_cq_tren.paragraph_format.space_after = Pt(2)
    p_cq_tren.paragraph_format.line_spacing = 1.15
    run1 = p_cq_tren.add_run(data.get("agency_upper", "ĐẢNG BỘ THÀNH PHỐ ĐÀ NẴNG"))
    run1.font.name = FONT_NAME
    run1.font.size = Pt(14)
    
    p_cq = cell_left.add_paragraph()
    p_cq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cq.paragraph_format.space_before = Pt(0)
    p_cq.paragraph_format.space_after = Pt(2)
    p_cq.paragraph_format.line_spacing = 1.15
    run2 = p_cq.add_run(data.get("agency_name", "ĐẢNG UỶ XÃ QUẾ SƠN"))
    run2.font.name = FONT_NAME
    run2.font.size = Pt(14)
    run2.font.bold = True
    
    p_star = cell_left.add_paragraph()
    p_star.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_star.paragraph_format.space_before = Pt(0)
    p_star.paragraph_format.space_after = Pt(6)
    run_star = p_star.add_run("*")
    run_star.font.name = FONT_NAME
    run_star.font.size = Pt(14)
    run_star.font.bold = True
    
    # Cột phải (Tiêu đề Quốc hiệu & ngày tháng theo đúng mẫu)
    p_dang = cell_right.paragraphs[0]
    p_dang.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dang.paragraph_format.space_before = Pt(0)
    p_dang.paragraph_format.space_after = Pt(2)
    p_dang.paragraph_format.line_spacing = 1.15
    run3 = p_dang.add_run("ĐẢNG CỘNG SẢN VIỆT NAM")
    run3.font.name = FONT_NAME
    run3.font.size = Pt(15)
    run3.font.bold = True
    run3.font.underline = True
    
    p_date = cell_right.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date.paragraph_format.space_before = Pt(4)
    p_date.paragraph_format.space_after = Pt(6)
    
    # Chuẩn hóa ngày tháng: đảm bảo ngày có số 0 đằng trước nếu < 10 (ví dụ ngày 01 tháng 11 năm 2026)
    raw_date_text = data.get("doc_date_text", "Quế Sơn, ngày 01 tháng 11 năm 2026")
    import re
    norm_date_text = re.sub(r'ngày\s+(\d)\b', r'ngày 0\1', raw_date_text)
    
    run4 = p_date.add_run(norm_date_text)
    run4.font.name = FONT_NAME
    run4.font.size = Pt(14)
    run4.font.italic = True
    
    # 3. Tiêu đề Thông báo
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(1)
    run_t1 = p_title.add_run("THÔNG BÁO")
    run_t1.font.name = FONT_NAME
    run_t1.font.size = Pt(15)
    run_t1.font.bold = True
    
    p_subtitle = doc.add_paragraph()
    p_subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_subtitle.paragraph_format.space_before = Pt(0)
    p_subtitle.paragraph_format.space_after = Pt(1)
    run_t2 = p_subtitle.add_run(data.get("doc_title", "lịch công tác của Ban Thường vụ Đảng ủy từ ngày 26/10/2026 đến ngày 01/11/2026"))
    run_t2.font.name = FONT_NAME
    run_t2.font.size = Pt(13)
    run_t2.font.bold = True
    
    p_dash = doc.add_paragraph()
    p_dash.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dash.paragraph_format.space_before = Pt(0)
    p_dash.paragraph_format.space_after = Pt(4)
    run_d = p_dash.add_run("-----")
    run_d.font.name = FONT_NAME
    run_d.font.size = Pt(12)
    run_d.font.bold = True
    
    # 4. Bảng Lịch Công Tác
    days_data = data.get("days", [])
    
    # Lọc Thứ 7 và Chủ nhật: Nếu không có nội dung công việc thì KHÔNG hiển thị trên file word
    filtered_days = []
    for d in days_data:
        dname = d.get("day_name", "")
        is_weekend = any(w in dname.lower() for w in ["thứ bảy", "thứ 7", "chủ nhật"])
        m_events = d.get("morning", [])
        a_events = d.get("afternoon", [])
        has_content = any(bool(e.get("content", "").strip() or e.get("leader", "").strip()) for e in (m_events + a_events))
        if is_weekend and not has_content:
            continue
        filtered_days.append(d)
        
    flat_rows = []
    day_row_spans = [] # (start_idx, count)
    session_row_spans = [] # (start_idx, count)
    
    current_idx = 1 # hàng 0 là header
    for d in filtered_days:
        day_start = current_idx
        day_events_count = 0
        
        # Buổi sáng
        morning_events = d.get("morning", [])
        if morning_events:
            s_start = current_idx
            m_len = len(morning_events)
            for ev_i, ev in enumerate(morning_events):
                flat_rows.append({
                    "day_label": f"{d.get('day_name', '')}\n{d.get('date_str', '')}" if (day_events_count == 0) else "",
                    "session_label": "S" if (ev_i == 0) else "",
                    "leader": ev.get("leader", ""),
                    "content": ev.get("content", ""),
                    "location": ev.get("location", ""),
                    "staff": ev.get("staff", ""),
                    "is_red": ev.get("is_red", False),
                    "has_next_in_session": (ev_i < m_len - 1)
                })
                current_idx += 1
                day_events_count += 1
            session_row_spans.append((s_start, m_len))
            
        # Buổi chiều
        afternoon_events = d.get("afternoon", [])
        if afternoon_events:
            c_start = current_idx
            a_len = len(afternoon_events)
            for ev_i, ev in enumerate(afternoon_events):
                flat_rows.append({
                    "day_label": f"{d.get('day_name', '')}\n{d.get('date_str', '')}" if (day_events_count == 0) else "",
                    "session_label": "C" if (ev_i == 0) else "",
                    "leader": ev.get("leader", ""),
                    "content": ev.get("content", ""),
                    "location": ev.get("location", ""),
                    "staff": ev.get("staff", ""),
                    "is_red": ev.get("is_red", False),
                    "has_next_in_session": (ev_i < a_len - 1)
                })
                current_idx += 1
                day_events_count += 1
            session_row_spans.append((c_start, a_len))
            
        # Nếu cả sáng và chiều đều không có sự kiện nào thì tạo 1 hàng trống
        if day_events_count == 0:
            flat_rows.append({
                "day_label": f"{d.get('day_name', '')}\n{d.get('date_str', '')}",
                "session_label": "",
                "leader": "",
                "content": "",
                "location": "",
                "staff": "",
                "is_red": False,
                "has_next_in_session": False
            })
            day_events_count = 1
            current_idx += 1
            
        day_row_spans.append((day_start, day_events_count))
        
    total_table_rows = 1 + len(flat_rows)
    table = doc.add_table(rows=total_table_rows, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="000000", sz="4", val="single")
    
    col_widths = [Cm(2.2), Cm(0.7), Cm(2.7), Cm(7.7), Cm(2.7), Cm(2.3)]
    
    # Hàng 0: Header
    header_row = table.rows[0]
    make_row_header(header_row)
    make_row_cant_split(header_row)
    
    headers_text = ["Ngày", "", "TT-BTV ĐU", "Nội dung", "Địa điểm", "VP + LX"]
    for c_i in range(6):
        c = header_row.cells[c_i]
        c.width = col_widths[c_i]
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c, top=60, bottom=60, left=50, right=50)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        if headers_text[c_i]:
            run = p.add_run(headers_text[c_i])
            run.font.name = FONT_NAME
            run.font.size = Pt(13 if c_i == 0 else 12)
            run.font.bold = True
        # Đường viền đỉnh đôi thanh đậm và đường viền đáy đôi mảnh ngăn cách tiêu đề và thân bảng, viền dọc đôi
        set_cell_border(
            c,
            top={'val': 'thinThickLargeGap', 'sz': '24'},
            bottom={'val': 'double', 'sz': '4'},
            left={'val': 'thinThickLargeGap', 'sz': '24'} if c_i == 0 else {'val': 'double', 'sz': '4'},
            right={'val': 'thinThickLargeGap', 'sz': '24'} if c_i == 5 else {'val': 'double', 'sz': '4'}
        )
            
    header_row.cells[0].merge(header_row.cells[1])
    
    # Điền nội dung từng hàng dữ liệu
    for r_i, r_data in enumerate(flat_rows, start=1):
        row = table.rows[r_i]
        make_row_cant_split(row)
        is_red = r_data["is_red"]
        red_color = RGBColor(255, 0, 0) if is_red else RGBColor(0, 0, 0)
        
        # Col 0: Ngày
        c0 = row.cells[0]
        c0.width = col_widths[0]
        c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c0, top=40, bottom=40, left=40, right=40)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        p0.paragraph_format.line_spacing = 1.05
        if r_data["day_label"]:
            run0 = p0.add_run(r_data["day_label"])
            run0.font.name = FONT_NAME
            run0.font.size = Pt(12)
            run0.font.bold = True
        
        # Col 1: Buổi (S/C)
        c1 = row.cells[1]
        c1.width = col_widths[1]
        c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c1, top=40, bottom=40, left=20, right=20)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        if r_data["session_label"]:
            run1 = p1.add_run(r_data["session_label"])
            run1.font.name = FONT_NAME
            run1.font.size = Pt(12)
            run1.font.bold = True
        
        # Col 2: TT-BTV ĐU
        c2 = row.cells[2]
        c2.width = col_widths[2]
        c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c2, top=40, bottom=40, left=40, right=40)
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing = 1.05
        run2 = p2.add_run(r_data["leader"])
        run2.font.name = FONT_NAME
        run2.font.size = Pt(12)
        if is_red:
            run2.font.color.rgb = red_color
            
        # Col 3: Nội dung
        c3 = row.cells[3]
        c3.width = col_widths[3]
        c3.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c3, top=40, bottom=40, left=50, right=50)
        p3 = c3.paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(0)
        p3.paragraph_format.line_spacing = 1.05
        run3 = p3.add_run(r_data["content"])
        run3.font.name = FONT_NAME
        run3.font.size = Pt(12)
        if is_red:
            run3.font.color.rgb = red_color
            
        # Col 4: Địa điểm
        c4 = row.cells[4]
        c4.width = col_widths[4]
        c4.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c4, top=40, bottom=40, left=40, right=40)
        p4 = c4.paragraphs[0]
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p4.paragraph_format.space_before = Pt(0)
        p4.paragraph_format.space_after = Pt(0)
        p4.paragraph_format.line_spacing = 1.05
        run4 = p4.add_run(r_data["location"])
        run4.font.name = FONT_NAME
        run4.font.size = Pt(12)
        if is_red:
            run4.font.color.rgb = red_color
            
        # Col 5: VP + LX
        c5 = row.cells[5]
        c5.width = col_widths[5]
        c5.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c5, top=40, bottom=40, left=40, right=40)
        p5 = c5.paragraphs[0]
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p5.paragraph_format.space_before = Pt(0)
        p5.paragraph_format.space_after = Pt(0)
        p5.paragraph_format.line_spacing = 1.05
        run5 = p5.add_run(r_data["staff"])
        run5.font.name = FONT_NAME
        run5.font.size = Pt(11.5)
        if is_red:
            run5.font.color.rgb = red_color

        # Line ngăn cách và viền dọc kép cho các cột trong bảng (giống mẫu gốc)
        is_last_table_row = (r_i == len(flat_rows))
        for col_idx in range(6):
            left_b = {'val': 'thinThickLargeGap', 'sz': '24'} if col_idx == 0 else {'val': 'double', 'sz': '4'}
            right_b = {'val': 'thinThickLargeGap', 'sz': '24'} if col_idx == 5 else {'val': 'double', 'sz': '4'}
            
            # Viền ngang dưới (dotted nếu còn sự kiện cùng buổi, single nếu hết buổi/hết ngày)
            if r_data.get("has_next_in_session") and col_idx in [2, 3, 4, 5]:
                bot_b = {'val': 'dotted', 'sz': '4'}
                if r_i + 1 < total_table_rows:
                    set_cell_border(table.rows[r_i + 1].cells[col_idx], top={'val': 'dotted', 'sz': '4'})
            else:
                bot_b = {'val': 'single', 'sz': '4'}
                
            set_cell_border(row.cells[col_idx], left=left_b, right=right_b, bottom=bot_b)

    # Merge các ô theo chiều dọc:
    # 1. Merge Buổi (Col 1)
    for s_start, count in session_row_spans:
        if count > 1:
            top_cell = table.cell(s_start, 1)
            bottom_cell = table.cell(s_start + count - 1, 1)
            top_cell.merge(bottom_cell)
            top_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
    # 2. Merge Ngày (Col 0)
    for d_start, count in day_row_spans:
        if count > 1:
            top_cell = table.cell(d_start, 0)
            bottom_cell = table.cell(d_start + count - 1, 0)
            top_cell.merge(bottom_cell)
            top_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # 5. Ghi chú sau bảng
    p_note1 = doc.add_paragraph()
    p_note1.paragraph_format.space_before = Pt(8)
    p_note1.paragraph_format.space_after = Pt(2)
    p_note1.paragraph_format.line_spacing = 1.15
    run_n1 = p_note1.add_run(data.get("driver_duty_note", "* Thứ Bảy, Chủ nhật, ngày 31/10 – 01/11/2026. A.Lực trực lái xe Đảng ủy."))
    run_n1.font.name = FONT_NAME
    run_n1.font.size = Pt(12)
    run_n1.font.italic = True
    
    p_note2 = doc.add_paragraph()
    p_note2.paragraph_format.space_before = Pt(0)
    p_note2.paragraph_format.space_after = Pt(4)
    p_note2.paragraph_format.line_spacing = 1.15
    run_n2 = p_note2.add_run(data.get("invitation_note", "* Lịch công tác thay cho giấy mời họp Thường trực, Ban Thường vụ Đảng ủy."))
    run_n2.font.name = FONT_NAME
    run_n2.font.size = Pt(12)
    run_n2.font.italic = True
    
    # Ghi chú giải thích chữ màu đỏ (nếu có)
    change_note = data.get("change_note")
    if change_note:
        p_c_note = doc.add_paragraph()
        p_c_note.paragraph_format.space_before = Pt(0)
        p_c_note.paragraph_format.space_after = Pt(4)
        p_c_note.paragraph_format.line_spacing = 1.15
        run_cn = p_c_note.add_run(change_note)
        run_cn.font.name = FONT_NAME
        run_cn.font.size = Pt(12)
        run_cn.font.italic = True
        run_cn.font.color.rgb = RGBColor(255, 0, 0)
    
    # Dự kiến tuần sau
    next_week_title = data.get("next_week_title", "* Tuần 45:")
    if next_week_title:
        p_nw = doc.add_paragraph()
        p_nw.paragraph_format.space_before = Pt(4)
        p_nw.paragraph_format.space_after = Pt(2)
        run_nw = p_nw.add_run(next_week_title)
        run_nw.font.name = FONT_NAME
        run_nw.font.size = Pt(12)
        run_nw.font.bold = True
        run_nw.font.italic = True
        
    next_week_items = data.get("next_week_items", [])
    for item in next_week_items:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.space_before = Pt(0)
        p_item.paragraph_format.space_after = Pt(3)
        p_item.paragraph_format.line_spacing = 1.15
        p_item.paragraph_format.left_indent = Inches(0.15)
        run_it = p_item.add_run(item)
        run_it.font.name = FONT_NAME
        run_it.font.size = Pt(12)
        run_it.font.italic = True
        
    # 6. Phần Nơi nhận & Thẩm quyền Ký
    sign_table = doc.add_table(rows=1, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_table.autofit = False
    
    c_recipients = sign_table.cell(0, 0)
    c_sign = sign_table.cell(0, 1)
    c_recipients.width = Cm(8.5)
    c_sign.width = Cm(9.9)
    
    # Cột Nơi nhận (Không in đậm, các dòng con ngang hàng lề trái)
    p_rec_title = c_recipients.paragraphs[0]
    p_rec_title.paragraph_format.space_before = Pt(14)
    p_rec_title.paragraph_format.space_after = Pt(2)
    p_rec_title.paragraph_format.left_indent = Pt(0)
    run_rt = p_rec_title.add_run("Nơi nhận:")
    run_rt.font.name = FONT_NAME
    run_rt.font.size = Pt(14)
    run_rt.font.bold = False
    run_rt.font.underline = True
    
    recipients = data.get("recipients", [
        "Ủy viên BTV Đảng ủy;",
        "Thường trực HĐND, UBND xã;",
        "Các cơ quan CTTM, GV Đảng ủy;",
        "Lưu Văn phòng Đảng ủy."
    ])
    for rec in recipients:
        p_r = c_recipients.add_paragraph()
        p_r.paragraph_format.space_before = Pt(0)
        p_r.paragraph_format.space_after = Pt(1)
        p_r.paragraph_format.line_spacing = 1.15
        p_r.paragraph_format.left_indent = Pt(0)
        run_r = p_r.add_run(f"- {rec}" if not rec.startswith("-") else rec)
        run_r.font.name = FONT_NAME
        run_r.font.size = Pt(12)
        
    # Cột Chữ ký
    p_auth = c_sign.paragraphs[0]
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_auth.paragraph_format.space_before = Pt(14)
    p_auth.paragraph_format.space_after = Pt(2)
    run_a1 = p_auth.add_run(data.get("signer_authority", "T/L BAN THƯỜNG VỤ"))
    run_a1.font.name = FONT_NAME
    run_a1.font.size = Pt(14)
    run_a1.font.bold = True
    
    p_role = c_sign.add_paragraph()
    p_role.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_role.paragraph_format.space_before = Pt(0)
    p_role.paragraph_format.space_after = Pt(4)
    run_a2 = p_role.add_run(data.get("signer_title", "CHÁNH VĂN PHÒNG"))
    run_a2.font.name = FONT_NAME
    run_a2.font.size = Pt(14)
    run_a2.font.bold = False
    
    # Chèn ảnh chữ ký nếu có
    sig_file = signature_path or data.get("signature_image")
    if sig_file and os.path.exists(sig_file) and data.get("include_signature", True):
        p_sig_img = c_sign.add_paragraph()
        p_sig_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sig_img.paragraph_format.space_before = Pt(2)
        p_sig_img.paragraph_format.space_after = Pt(2)
        p_sig_img.add_run().add_picture(sig_file, width=Inches(1.8))
    else:
        # Nếu không có ảnh chữ ký thì tạo khoảng cách để ký tay
        # p_space = c_sign.add_paragraph()
        # p_space.paragraph_format.space_before = Pt(40)
        for _ in range(4):
            c_sign.add_paragraph()
        
    p_name = c_sign.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(2)
    p_name.paragraph_format.space_after = Pt(0)
    run_name = p_name.add_run(data.get("signer_name", "Nguyễn Phước Tâm"))
    run_name.font.name = FONT_NAME
    run_name.font.size = Pt(14)
    run_name.font.bold = True
    
    doc.save(output_path)
    return output_path
