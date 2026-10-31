import urllib.request, urllib.parse, json, docx, sys, os, fitz, win32com.client
from lxml import etree
sys.stdout.reconfigure(encoding='utf-8')

# 1. Test API export
url = 'http://localhost:5000/api/sample-data'
req = urllib.request.Request(url)
with urllib.request.urlopen(req) as resp:
    sample_data = json.loads(resp.read().decode('utf-8'))

# Modify for Week 45
sample_data['doc_title'] = 'lịch công tác của Ban Thường vụ Đảng ủy từ ngày 02/11/2026 đến ngày 08/11/2026'
sample_data['next_week_title'] = '* Tuần 46:'
sample_data['doc_date_text'] = 'Quế Sơn, ngày 1 tháng 11 năm 2026'

export_url = 'http://localhost:5000/api/export-docx'
req2 = urllib.request.Request(export_url, data=json.dumps(sample_data).encode('utf-8'), headers={'Content-Type': 'application/json'})
with urllib.request.urlopen(req2) as resp2:
    x_filename = resp2.headers.get('X-Download-Filename')
    decoded_filename = urllib.parse.unquote(x_filename)
    print('Exported Filename Header:', decoded_filename)
    assert decoded_filename == 'Tuần 45 lịch công tác BTV 02.11 - 08.11.2026.docx', f'Filename mismatch: {decoded_filename}'
    
    out_docx = r'D:\PM\LichCongTacTuan\data\test_final_week45.docx'
    with open(out_docx, 'wb') as f:
        f.write(resp2.read())

print('Test 1 (Filename): PASS')

# 2. Check docx structure
doc = docx.Document(out_docx)
# Check table 0 (Header block)
t0 = doc.tables[0]
r_cell = t0.rows[0].cells[1]
p_dang = r_cell.paragraphs[0]
print('P_dang text:', p_dang.text, 'underline:', p_dang.runs[0].font.underline)
assert p_dang.runs[0].font.underline is None, 'Error: Underline must not be on the text run'

# Check table 1 (Schedule table)
t1 = doc.tables[1]
header_c0 = t1.rows[0].cells[0]
tcB = header_c0._tc.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders')
tcB_xml = etree.tostring(tcB).decode('utf-8')
print('Header bottom border in table 1:', [line for line in tcB_xml.splitlines() if 'bottom' in line])
assert 'w:val="double" w:sz="4"' in tcB_xml, 'Error: Double line must be sz=4 (thin)'
print('Test 2 (Double line thin sz=4): PASS')

# 3. Convert to PDF and PNG for visual inspection
out_pdf = r'D:\PM\LichCongTacTuan\data\test_final_week45.pdf'
out_png = r'D:\PM\LichCongTacTuan\data\test_final_week45.png'

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
d = word.Documents.Open(os.path.abspath(out_docx))
d.SaveAs2(os.path.abspath(out_pdf), FileFormat=17)
d.Close()
word.Quit()

doc_pdf = fitz.open(out_pdf)
page = doc_pdf.load_page(0)
pix = page.get_pixmap(dpi=150)
pix.save(out_png)
print('Test 3 (Visual PNG generated): PASS ->', out_png)
