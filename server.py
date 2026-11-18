import os
import sys
import json
from datetime import datetime, timedelta

# Cấu hình UTF-8 cho Windows console
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

from flask import Flask, render_template, request, send_file, send_from_directory, jsonify

# Import docx generator
from docx_generator import generate_weekly_schedule_docx
from word_parser import parse_weekly_schedule_file

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
HISTORY_DIR = os.path.join(DATA_DIR, "history")
STATIC_DIR = os.path.join(BASE_DIR, "static")
DOWNLOADS_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(HISTORY_DIR, exist_ok=True)
os.makedirs(DOWNLOADS_DIR, exist_ok=True)

DEFAULT_SIGNATURE_PATH = os.path.join(STATIC_DIR, "signature.png")

SAMPLE_DATA = {
    "agency_upper": "ĐẢNG BỘ THÀNH PHỐ ĐÀ NẴNG",
    "agency_name": "ĐẢNG UỶ XÃ QUẾ SƠN",
    "doc_date_text": "Quế Sơn, ngày 25 tháng 10 năm 2026",
    "doc_title": "lịch công tác của Ban Thường vụ Đảng ủy từ ngày 26/10/2026 đến ngày 01/11/2026",
    "days": [
        {
            "day_name": "Thứ Hai",
            "date_str": "26/10",
            "morning": [
                {"leader": "TT.ĐU,\nC.Hằng", "content": "- 08h00: Họp Thường trực Đảng ủy.", "location": "Phòng họp số 1", "staff": "A.Tâm", "is_red": False},
                {"leader": "A.Trung,\nA.Hùng", "content": "- 09h00: Họp đánh giá hồ sơ OCOP.", "location": "Phòng họp số 4", "staff": "", "is_red": False},
                {"leader": "A.Tùng", "content": "- 07h00: Hội nghị trực tuyến về quản lý tài sản công và chỉ tiêu thống kê GRDP.", "location": "Phòng họp số 2", "staff": "", "is_red": False}
            ],
            "afternoon": [
                {"leader": "A.Tùng", "content": "- 15h00: Làm việc với Sở Công thương về Tổ chức Chương trình Kết nối giao thương giữa doanh nghiệp, hợp tác xã, các cơ sở sản xuất, kinh doanh về tổ chức xúc tiến thương mại tại thành phố Đà Nẵng năm 2026.", "location": "Phòng họp số 2", "staff": "", "is_red": False}
            ]
        },
        {
            "day_name": "Thứ Ba",
            "date_str": "27/10",
            "morning": [
                {"leader": "A.Hùng,\nA.Công", "content": "- 08h00: Tham gia lớp bồi dưỡng cấp ủy cấp xã, phường năm 2026 (thời gian từ ngày 27/10 – 02/11/2026).", "location": "Số 18 Lam Sơn, P. Bàn Thạch", "staff": "", "is_red": False}
            ],
            "afternoon": [
                {"leader": "BTV", "content": "- 14h00: Họp Ban Thường vụ Đảng ủy.", "location": "Phòng họp số 1", "staff": "A.Tâm", "is_red": True},
                {"leader": "A.Tùng", "content": "- 14h00: Làm việc với các hộ dân thôn Châu Sơn về công tác môi trường.", "location": "Phòng họp số 2", "staff": "", "is_red": False}
            ]
        },
        {
            "day_name": "Thứ Tư",
            "date_str": "28/10",
            "morning": [
                {"leader": "BTV", "content": "- 08h00: Họp Ban Chỉ đạo cơ sở dữ liệu đất đai.", "location": "Phòng họp số 4", "staff": "C.Nguyệt", "is_red": False}
            ],
            "afternoon": [
                {"leader": "A.Vũ", "content": "- 14h00: Đi kiểm tra, khảo sát thực tế.", "location": "Tại thực tế", "staff": "A.Lực", "is_red": False},
                {"leader": "A.Trung,\nA.Tùng", "content": "- 14h00: Họp giao ban lãnh đạo UBND xã.", "location": "Phòng họp số 1", "staff": "", "is_red": False}
            ]
        },
        {
            "day_name": "Thứ Năm",
            "date_str": "29/10",
            "morning": [
                {"leader": "A. Vũ,\nA.Trung(QS)", "content": "- 08h00: Sinh hoạt Đảng bộ Quân sự.", "location": "Phòng họp BCHQS xã", "staff": "", "is_red": False},
                {"leader": "A.Trung,\nA.Hùng", "content": "- 08h00: Họp xét khuyết tật.", "location": "Hội trường xã Quế Sơn", "staff": "", "is_red": False}
            ],
            "afternoon": [
                {"leader": "TT.ĐU", "content": "- 14h00: Thường trực Đảng ủy làm việc với các Đảng bộ trường học trên địa bàn xã.", "location": "Phòng họp số 2", "staff": "A.Lưu,\nA.Phương", "is_red": False},
                {"leader": "A.Tùng", "content": "- 14h00: Làm việc với các hộ dân Tổ 16, thôn Đại Lộc.", "location": "Phòng họp số 1", "staff": "", "is_red": False}
            ]
        },
        {
            "day_name": "Thứ Sáu",
            "date_str": "30/10",
            "morning": [
                {"leader": "BTV", "content": "- 08h00: Họp HĐND xã lần thứ 4", "location": "Phòng họp số 5", "staff": "", "is_red": False}
            ],
            "afternoon": [
                {"leader": "A.Trung,\nA.Tùng", "content": "- 14h00: Dự nghe Trung tâm CUDVSN công xã báo cáo Phương án chỉnh trang khu Trung tâm xã", "location": "Phòng họp số 2", "staff": "", "is_red": True}
            ]
        }
    ],
    "driver_duty_note": "* Thứ Bảy, Chủ nhật, ngày 31/10 – 01/11/2026. A.Lực trực lái xe Đảng ủy.",
    "invitation_note": "* Lịch công tác thay cho giấy mời họp Thường trực, Ban Thường vụ Đảng ủy.",
    "next_week_title": "* Tuần 45:",
    "next_week_items": [
        "- 14h00 ngày 02/11/2026: BTV: Dự hội nghị quán triệt, triển khai các văn bản của Trung ương, Thành ủy tại hội trường xã Quế Sơn.",
        "- 14h00 ngày 03/11/2026: TTĐU: Họp tổ chức đánh giá việc thực hiện các Thông báo kết luận và các nội dung chỉ đạo của Trưởng Ban chỉ đạo tại các cuộc họp tại phòng họp số 2, xã Quế Sơn.",
        "- 08h00 ngày 04/11/2026: A.Vũ: Tổ đại biểu bầu cử số 12 của HĐND thành phố về tiếp xúc cử tri kỳ họp thường lệ năm 2026, khóa XI, nhiệm kỳ 2026-2031 tại số 545 Hùng Vương, xã Nam Phước.",
        "- 08h00 ngày 05/11/2026: A.Vũ: Tổ đại biểu bầu cử số 12 của HĐND thành phố về tiếp xúc cử tri kỳ họp thường lệ năm 2026, khóa XI, nhiệm kỳ 2026-2031 tại Hội trường xã Quế Sơn."
    ],
    "recipients": [
        "Ủy viên BTV Đảng ủy;",
        "Thường trực HĐND, UBND xã;",
        "Các cơ quan CTTM, GV Đảng ủy;",
        "Lưu Văn phòng Đảng ủy."
    ],
    "signer_authority": "T/L BAN THƯỜNG VỤ",
    "signer_title": "CHÁNH VĂN PHÒNG",
    "signer_name": "Nguyễn Phước Tâm",
    "include_signature": True
}

@app.route("/")
def index():
    """Trang chủ Landing Page OfficeFormat"""
    return render_template("landing.html")

@app.route("/lich-tuan")
@app.route("/app")
def schedule_app():
    """Ứng dụng Lập & Xuất Lịch Công Tác Tuần (Chuẩn Đảng & Nhà Nước)"""
    return render_template("index.html")

@app.route("/downloads/<path:filename>")
def download_file(filename):
    """Tải bộ cài đặt OfficeFormat"""
    return send_from_directory(DOWNLOADS_DIR, filename, as_attachment=True)

@app.route("/favicon.ico")
def favicon():
    """Favicon OfficeFormat"""
    return send_from_directory(STATIC_DIR, "favicon.ico", mimetype="image/vnd.microsoft.icon")

@app.route("/robots.txt")
def robots():
    """Robots SEO"""
    return send_from_directory(BASE_DIR, "robots.txt", mimetype="text/plain")

@app.route("/sitemap.xml")
def sitemap():
    """Sitemap SEO"""
    return send_from_directory(BASE_DIR, "sitemap.xml", mimetype="application/xml")

@app.route("/api/sample-data", methods=["GET"])
def get_sample_data():
    return jsonify(SAMPLE_DATA)

@app.route("/api/export-docx", methods=["POST"])
def export_docx():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dữ liệu trống"}), 400
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Thong_bao_lich_tuan_{timestamp}.docx"
        output_path = os.path.join(DATA_DIR, filename)
        
        sig_path = DEFAULT_SIGNATURE_PATH if os.path.exists(DEFAULT_SIGNATURE_PATH) else None
        generate_weekly_schedule_docx(data, output_path, sig_path)
        
        # Tạo tên file xuất chuẩn theo yêu cầu: "Tuần 45 lịch công tác BTV 02.11 - 08.11.2026.docx"
        import urllib.parse
        import re
        
        doc_title = (data.get("doc_title") or "").strip()
        days = data.get("days", [])
        
        week_num = None
        next_week_title = data.get("next_week_title", "")
        m_nw = re.search(r"Tuần\s*(\d+)", next_week_title, re.IGNORECASE)
        if m_nw:
            week_num = int(m_nw.group(1)) - 1
            
        m_dates = re.findall(r"(\d{1,2})[/-](\d{1,2})(?:[/-](\d{4}))?", doc_title)
        start_date_str = None
        end_date_str = None
        year_str = "2026"
        
        if len(m_dates) >= 2:
            d1, m1, y1 = m_dates[0]
            d2, m2, y2 = m_dates[1]
            year_str = y2 or y1 or str(datetime.now().year)
            start_date_str = f"{int(d1):02d}.{int(m1):02d}"
            end_date_str = f"{int(d2):02d}.{int(m2):02d}.{year_str}"
            if not week_num:
                try:
                    dt_start = datetime(int(year_str), int(m1), int(d1))
                    week_num = dt_start.isocalendar()[1]
                except Exception:
                    week_num = 45
        elif days and len(days) > 0:
            fd_parts = days[0].get("date_str", "").split("/")
            ld_parts = days[-1].get("date_str", "").split("/")
            if len(fd_parts) >= 2:
                start_date_str = f"{int(fd_parts[0]):02d}.{int(fd_parts[1]):02d}"
            if len(ld_parts) >= 2:
                end_date_str = f"{int(ld_parts[0]):02d}.{int(ld_parts[1]):02d}.{year_str}"
                
        if not week_num:
            week_num = 45
        if not start_date_str:
            start_date_str = "02.11"
        if not end_date_str:
            end_date_str = f"08.11.{year_str}"
            
        download_filename = f"Tuần {week_num} lịch công tác BTV {start_date_str} - {end_date_str}.docx"
        ascii_fallback = f"Tuan_{week_num}_lich_cong_tac_BTV_{start_date_str}_-_{end_date_str}.docx"
                
        response = send_file(
            output_path,
            as_attachment=True,
            download_name=ascii_fallback,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        response.headers['X-Download-Filename'] = urllib.parse.quote(download_filename)
        return response
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route("/api/parse-word", methods=["POST"])
def parse_word_upload():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "Vui lòng chọn file Word (.docx hoặc .doc)"}), 400
            
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "Chưa chọn file"}), 400
            
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ['.docx', '.doc']:
            return jsonify({"error": "Chỉ hỗ trợ file Word (.docx hoặc .doc)"}), 400
            
        # Lưu file tạm để phân tích
        temp_upload_path = os.path.join(DATA_DIR, f"upload_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}")
        file.save(temp_upload_path)
        
        try:
            parsed_data = parse_weekly_schedule_file(temp_upload_path)
            
            # Đặt tên lịch theo tiêu đề hoặc tên file
            clean_name = parsed_data.get("doc_title") or file.filename.rsplit('.', 1)[0]
            if "filename" in clean_name.lower() or "utf-8" in clean_name.lower() or ";" in clean_name:
                clean_name = parsed_data.get("doc_title") or "Lịch công tác tuần"
            parsed_data["name"] = clean_name
            parsed_data["saved_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Tự động lưu vào kho lịch cũ để dùng lại lần sau
            hist_id = f"word_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            parsed_data["id"] = hist_id
            fpath = os.path.join(HISTORY_DIR, f"{hist_id}.json")
            with open(fpath, "w", encoding="utf-8") as f:
                json.dump(parsed_data, f, ensure_ascii=False, indent=2)
                
            return jsonify({
                "success": True,
                "data": parsed_data,
                "filename": file.filename
            })
        finally:
            if os.path.exists(temp_upload_path):
                try: os.remove(temp_upload_path)
                except: pass
                
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Lỗi bóc tách file Word: {str(e)}"}), 500

@app.route("/api/history", methods=["GET"])
def get_history_list():
    try:
        items = []
        for fname in os.listdir(HISTORY_DIR):
            if fname.endswith(".json"):
                fpath = os.path.join(HISTORY_DIR, fname)
                try:
                    with open(fpath, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    
                    # Tính toán tổng số cuộc họp
                    total_events = 0
                    red_events = 0
                    for d in data.get("days", []):
                        for ev in d.get("morning", []):
                            total_events += 1
                            if ev.get("is_red"): red_events += 1
                        for ev in d.get("afternoon", []):
                            total_events += 1
                            if ev.get("is_red"): red_events += 1
                            
                    items.append({
                        "id": fname[:-5],
                        "name": data.get("name") or data.get("doc_title", fname[:-5]),
                        "doc_title": data.get("doc_title", ""),
                        "doc_date_text": data.get("doc_date_text", ""),
                        "saved_at": data.get("saved_at") or datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%Y-%m-%d %H:%M:%S"),
                        "total_events": total_events,
                        "red_events": red_events,
                        "file_size": os.path.getsize(fpath)
                    })
                except Exception as e:
                    print(f"Error reading {fname}: {e}")
                    
        # Sắp xếp lịch mới nhất lên đầu
        items.sort(key=lambda x: x["saved_at"], reverse=True)
        return jsonify(items)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/<item_id>", methods=["GET"])
def get_history_detail(item_id):
    try:
        safe_id = "".join(c for c in item_id if c.isalnum() or c in ("-", "_"))
        fpath = os.path.join(HISTORY_DIR, f"{safe_id}.json")
        if not os.path.exists(fpath):
            return jsonify({"error": "Không tìm thấy lịch cũ này"}), 404
        with open(fpath, "r", encoding="utf-8") as f:
            data = json.load(f)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/save", methods=["POST"])
def save_schedule_to_history():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "Dữ liệu trống"}), 400
            
        custom_name = data.get("save_name")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        timestamp_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Tạo ID an toàn
        item_id = f"lich_{timestamp_id}"
        if custom_name:
            data["name"] = custom_name
        else:
            data["name"] = data.get("doc_title", f"Lịch lưu lúc {now_str}")
            
        data["saved_at"] = now_str
        data["id"] = item_id
        
        fpath = os.path.join(HISTORY_DIR, f"{item_id}.json")
        with open(fpath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            
        return jsonify({"success": True, "id": item_id, "name": data["name"]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/<item_id>", methods=["DELETE"])
def delete_history_item(item_id):
    try:
        safe_id = "".join(c for c in item_id if c.isalnum() or c in ("-", "_"))
        fpath = os.path.join(HISTORY_DIR, f"{safe_id}.json")
        if os.path.exists(fpath):
            os.remove(fpath)
            return jsonify({"success": True})
        return jsonify({"error": "Không tìm thấy file"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 60)
    print(" ỨNG DỤNG LẬP & XUẤT LỊCH CÔNG TÁC TUẦN")
    print(f" Địa chỉ truy cập: http://0.0.0.0:{port}")
    print("=" * 60)
    app.run(host="0.0.0.0", port=port, debug=False)
