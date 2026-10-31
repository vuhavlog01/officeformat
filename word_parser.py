import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor

import sys

def convert_doc_to_docx(doc_path, out_docx_path):
    """Chuyển đổi file .doc cũ sang .docx bằng Microsoft Word COM (trên Windows) hoặc thông báo nếu trên Linux."""
    if sys.platform == 'win32':
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            try:
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False
                word.DisplayAlerts = False
                try:
                    doc = word.Documents.Open(os.path.abspath(doc_path))
                    doc.SaveAs2(os.path.abspath(out_docx_path), FileFormat=16) # wdFormatXMLDocument = 16
                    doc.Close(False)
                finally:
                    word.Quit()
            finally:
                pythoncom.CoUninitialize()
            return out_docx_path
        except Exception as e:
            pass
    raise RuntimeError("Máy chủ đám mây (Linux) không hỗ trợ định dạng .doc cũ. Vui lòng mở file bằng Word trên máy tính và lưu thành file .docx trước khi nạp!")

def is_binary_doc_file(file_path):
    """Kiểm tra xem file có phải định dạng Word nhị phân cũ .doc (OLE2 Compound Document) không."""
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
            return header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
    except Exception:
        return False

def parse_weekly_schedule_file(file_path):
    """
    Đọc và bóc tách toàn bộ dữ liệu lịch công tác tuần từ file Word (.docx hoặc .doc).
    Trả về cấu trúc dữ liệu JSON chuẩn để nạp vào giao diện Web App.
    """
    temp_docx = None
    is_doc = file_path.lower().endswith('.doc') or is_binary_doc_file(file_path)
    
    if is_doc:
        temp_docx = file_path + '.temp.docx'
        convert_doc_to_docx(file_path, temp_docx)
        actual_path = temp_docx
    else:
        actual_path = file_path

    try:
        try:
            doc = docx.Document(actual_path)
        except Exception:
            # Nếu mở trực tiếp bằng docx thất bại thì thử convert qua Word COM
            if not temp_docx:
                temp_docx = file_path + '.temp.docx'
                convert_doc_to_docx(file_path, temp_docx)
                actual_path = temp_docx
                doc = docx.Document(actual_path)
            else:
                raise
        
        result = {
            "agency_upper": "ĐẢNG BỘ THÀNH PHỐ ĐÀ NẴNG",
            "agency_name": "ĐẢNG UỶ XÃ QUẾ SƠN",
            "doc_date_text": "",
            "doc_title": "",
            "days": [],
            "driver_duty_note": "",
            "invitation_note": "",
            "change_note": "",
            "next_week_title": "",
            "next_week_items": [],
            "recipients": [],
            "signer_authority": "T/L BAN THƯỜNG VỤ",
            "signer_title": "CHÁNH VĂN PHÒNG",
            "signer_name": "Nguyễn Phước Tâm",
            "include_signature": True
        }
        
        # 1. Tìm bảng Lịch Công Tác
        sched_tbl = None
        sched_tbl_idx = -1
        for t_idx, t in enumerate(doc.tables):
            header_text = ' '.join(c.text for c in t.rows[0].cells).lower()
            if 'ngày' in header_text and ('nội dung' in header_text or 'tt-btv' in header_text):
                sched_tbl = t
                sched_tbl_idx = t_idx
                break
                
        # 2. Bóc tách tiêu đề, ngày tháng, cơ quan
        # Kiểm tra bảng Header (nếu có bảng 2 cột ở đầu)
        if sched_tbl_idx > 0:
            h_table = doc.tables[0]
            if len(h_table.rows) > 0 and len(h_table.rows[0].cells) >= 2:
                c_left = h_table.cell(0, 0).text.strip().split('\n')
                if len(c_left) >= 1 and c_left[0].strip():
                    result['agency_upper'] = c_left[0].strip()
                if len(c_left) >= 2 and c_left[1].strip():
                    result['agency_name'] = c_left[1].strip()
                    
                c_right = h_table.cell(0, 1).text.strip().split('\n')
                for line in c_right:
                    if 'ngày' in line.lower() and 'tháng' in line.lower():
                        result['doc_date_text'] = line.strip()
                        
        # Quét các đoạn văn bản (paragraphs)
        in_next_week = False
        in_recipients = False
        
        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue
                
            # Đơn vị cấp trên
            if 'ĐẢNG BỘ' in txt.upper() and len(txt) < 60 and not result.get('agency_upper_found'):
                # Tách dòng nếu có nhiều tab
                parts = txt.split()
                if 'ĐÀ NẴNG' in txt.upper() or 'TỈNH' in txt.upper() or 'THÀNH PHỐ' in txt.upper():
                    result['agency_upper'] = txt.split('    ')[0].strip()
                    result['agency_upper_found'] = True
                    
            # Tên đơn vị ban hành
            if 'ĐẢNG UỶ' in txt.upper() and len(txt) < 60 and not result.get('agency_name_found'):
                result['agency_name'] = txt.split('    ')[0].strip()
                result['agency_name_found'] = True
                
            # Ngày ban hành
            if ('ngày' in txt.lower() and 'tháng' in txt.lower() and 'năm' in txt.lower()) and not result['doc_date_text']:
                if 'Quế Sơn' in txt or ',' in txt:
                    # Lấy phần sau dấu tab hoặc dấu sao
                    match = re.search(r'([\w\s]+,\s*ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+)', txt)
                    if match:
                        result['doc_date_text'] = match.group(1).strip()
                    else:
                        result['doc_date_text'] = txt.split('\t')[-1].replace('*', '').strip()
                        
            # Tiêu đề thông báo
            if 'lịch công tác của Ban Thường vụ' in txt or ('lịch công tác' in txt.lower() and 'từ ngày' in txt.lower()):
                result['doc_title'] = txt
                
            # Ghi chú trực lái xe
            if 'trực lái xe' in txt.lower():
                result['driver_duty_note'] = txt
                
            # Ghi chú thay giấy mời
            if 'thay cho giấy mời' in txt.lower():
                result['invitation_note'] = txt
                
            # Ghi chú giải thích chữ màu đỏ
            if 'chữ màu đỏ' in txt.lower() or 'chữ in màu đỏ' in txt.lower():
                result['change_note'] = txt
                
            # Dự kiến lịch tuần sau
            if txt.startswith('* Tuần') or txt.startswith('Tuần '):
                result['next_week_title'] = txt
                in_next_week = True
                continue
                
            if in_next_week:
                if txt.startswith('-') or txt.startswith('•') or txt.startswith('+'):
                    result['next_week_items'].append(txt)
                elif 'Nơi nhận:' in txt or 'T/L' in txt:
                    in_next_week = False
                    
            # Nơi nhận trong paragraphs (nếu không dùng bảng)
            if 'Nơi nhận:' in txt:
                in_recipients = True
                # Tách phần thẩm quyền ký nếu nằm cùng dòng qua khoảng trắng lớn
                if 'T/L' in txt or 'T/M' in txt:
                    parts = re.split(r'\s{4,}|\t', txt)
                    for pt in parts:
                        if 'T/L' in pt or 'T/M' in pt:
                            result['signer_authority'] = pt.strip()
                continue
                
            if in_recipients:
                if 'CHÁNH VĂN PHÒNG' in txt or 'BÍ THƯ' in txt or 'PHÓ BÍ THƯ' in txt:
                    parts = re.split(r'\s{4,}|\t', txt)
                    for pt in parts:
                        if any(k in pt for k in ['CHÁNH VĂN PHÒNG', 'BÍ THƯ', 'PHÓ BÍ THƯ']):
                            result['signer_title'] = pt.strip()
                        elif pt.startswith('-'):
                            result['recipients'].append(pt.strip())
                elif any(k in txt for k in ['Nguyễn', 'Trần', 'Lê', 'Phạm', 'Huỳnh', 'Hoàng', 'Võ', 'Phan']) and len(txt) < 50 and not txt.startswith('-'):
                    result['signer_name'] = txt.strip()
                    in_recipients = False
                elif txt.startswith('-'):
                    # Tách phần chức danh / tên người ký nếu cùng dòng
                    parts = re.split(r'\s{4,}|\t', txt)
                    result['recipients'].append(parts[0].strip())
                    if len(parts) > 1:
                        p_right = parts[-1].strip()
                        if 'T/L' in p_right or 'T/M' in p_right:
                            result['signer_authority'] = p_right
                        elif 'CHÁNH VĂN PHÒNG' in p_right or 'BÍ THƯ' in p_right:
                            result['signer_title'] = p_right
                        elif len(p_right) > 3 and not p_right.startswith('-'):
                            result['signer_name'] = p_right
                elif 'Lưu' in txt:
                    result['recipients'].append(txt.strip())
                    
        # 3. Kiểm tra bảng Chữ ký & Nơi nhận (nếu có bảng riêng ở cuối)
        for t_idx, t in enumerate(doc.tables):
            if t_idx == sched_tbl_idx:
                continue
            t_text = ' '.join(c.text for c in t.rows[0].cells)
            if 'Nơi nhận' in t_text or 'CHÁNH VĂN PHÒNG' in t_text or 'T/L' in t_text or 'T/M' in t_text:
                rec_list = []
                for row in t.rows:
                    if len(row.cells) >= 2:
                        c_left_lines = row.cells[0].text.strip().split('\n')
                        c_right_lines = row.cells[1].text.strip().split('\n')
                        for line in c_left_lines:
                            line = line.strip()
                            if line and 'Nơi nhận' not in line:
                                rec_list.append(line)
                        for line in c_right_lines:
                            line = line.strip()
                            if 'T/L' in line or 'T/M' in line:
                                result['signer_authority'] = line
                            elif 'CHÁNH VĂN PHÒNG' in line or 'BÍ THƯ' in line:
                                result['signer_title'] = line
                            elif len(line) > 3 and not any(k in line for k in ['T/L', 'CHÁNH', 'VĂN PHÒNG', 'BÍ THƯ']):
                                result['signer_name'] = line
                if rec_list:
                    result['recipients'] = rec_list

        # 4. Bóc tách bảng Lịch Công Tác
        if sched_tbl:
            current_day = None
            current_session = 'morning'
            days = []
            day_map = {}
            
            for r_idx in range(1, len(sched_tbl.rows)):
                row = sched_tbl.rows[r_idx]
                unique_cells = []
                seen = []
                for c in row.cells:
                    if c._tc not in seen:
                        seen.append(c._tc)
                        unique_cells.append(c)
                        
                c0_text = unique_cells[0].text.strip()
                c1_text = unique_cells[1].text.strip() if len(unique_cells) > 1 else ''
                
                # Xác định ngày
                for dname in ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy', 'Chủ Nhật']:
                    if dname in c0_text:
                        current_day = dname
                        date_str = ''
                        for p in c0_text.split():
                            if '/' in p:
                                date_str = p.strip()
                        if current_day not in day_map:
                            day_obj = {'day_name': current_day, 'date_str': date_str, 'morning': [], 'afternoon': []}
                            day_map[current_day] = day_obj
                            days.append(day_obj)
                        break
                        
                # Xác định buổi
                if 'S' in c1_text:
                    current_session = 'morning'
                elif 'C' in c1_text:
                    current_session = 'afternoon'
                    
                # Phân bổ các cột dữ liệu
                if len(unique_cells) == 6:
                    c_leader, c_content, c_loc, c_staff = unique_cells[2], unique_cells[3], unique_cells[4], unique_cells[5]
                elif len(unique_cells) == 5:
                    c_leader, c_content, c_loc, c_staff = unique_cells[1], unique_cells[2], unique_cells[3], unique_cells[4]
                elif len(unique_cells) == 4:
                    c_leader, c_content, c_loc, c_staff = unique_cells[0], unique_cells[1], unique_cells[2], unique_cells[3]
                else:
                    continue
                    
                leader = c_leader.text.strip()
                content = c_content.text.strip()
                location = c_loc.text.strip()
                staff = c_staff.text.strip()
                
                # Nhận diện chữ ĐỎ (lịch thay đổi / điều chỉnh)
                is_red = False
                for cell in [c_leader, c_content, c_loc, c_staff]:
                    xml_str = cell._tc.xml
                    if 'FF0000' in xml_str or 'ff0000' in xml_str or 'color w:val="red"' in xml_str.lower():
                        is_red = True
                        break
                        
                if current_day in day_map and (leader or content):
                    day_map[current_day][current_session].append({
                        'leader': leader,
                        'content': content,
                        'location': location,
                        'staff': staff,
                        'is_red': is_red
                    })
                    
            if days:
                result['days'] = days

        # Dọn dẹp trường tạm
        result.pop('agency_upper_found', None)
        result.pop('agency_name_found', None)
        
        # Mặc định chuẩn nếu thiếu
        if not result['agency_upper']: result['agency_upper'] = "ĐẢNG BỘ THÀNH PHỐ ĐÀ NẴNG"
        if not result['agency_name']: result['agency_name'] = "ĐẢNG UỶ XÃ QUẾ SƠN"
        if not result['signer_name']: result['signer_name'] = "Nguyễn Phước Tâm"
        if not result['recipients']:
            result['recipients'] = [
                "Ủy viên BTV Đảng ủy;",
                "Thường trực HĐND, UBND xã;",
                "Các cơ quan CTTM, GV Đảng ủy;",
                "Lưu Văn phòng Đảng ủy."
            ]

        return result
    finally:
        if temp_docx and os.path.exists(temp_docx):
            try: os.remove(temp_docx)
            except: pass
