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
sample_data['doc_date_text'] = 'Quế Sơn, ngày 1 tháng 11 năm 2026' # Should be normalized to 01

export_url = 'http://localhost:5000/api/export-docx'
req2 = urllib.request.Request(export_url, data=json.dumps(sample_data).encode('utf-8'), headers={'Content-Type': 'application/json'})
with urllib.request.urlopen(req2) as resp2:
    x_filename = resp2.headers.get('X-Download-Filename')
    decoded_filename = urllib.parse.unquote(x_filename)
    print('Exported Filename Header:', decoded_filename)
    assert decoded_filename == 'Tuần 45 lịch công tác BTV 02.11 - 08.11.2026.docx', f'Filename mismatch: {decoded_filename}'
    
    out_docx = r'D:\PM\LichCongTacTuan\data\test_all_reqs.docx'
    with open(out_docx, 'wb') as f:
        f.write(resp2.read())

print('Test 1 (Filename): PASS')

# 2. Check docx structure
doc = docx.Document(out_docx)

# Check table 0 (Header block)
t0 = doc.tables[0]
r_cell = t0.rows[0].cells[1]
p_dang = r_cell.paragraphs[0]
print('P_dang text:', repr(p_dang.text), 'underline:', p_dang.runs[0].font.underline, 'align:', p_dang.alignment)
assert p_dang.runs[0].font.underline is True, 'Error: Underline must be on ĐẢNG CỘNG SẢN VIỆT NAM'

p_date = r_cell.paragraphs[1]
print('P_date text:', repr(p_date.text), 'align:', p_date.alignment)
assert 'ngày 01 tháng' in p_date.text, 'Error: Date must have leading zero 01'
assert p_date.alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER, 'Error: Date must be center aligned'
print('Test 2 (Quoc hieu tieu ngu giong mau): PASS')

# Check table 1 (Schedule table borders)
t1 = doc.tables[1]
# Header borders
header_c0 = t1.rows[0].cells[0]
tcB = header_c0._tc.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders')
tcB_xml = etree.tostring(tcB).decode('utf-8')
assert 'w:bottom w:val="double" w:sz="4"' in tcB_xml, 'Error: Header bottom must be double sz=4'

# Check vertical borders between cells in row 1
r1 = t1.rows[1]
for c_idx in range(len(r1.cells)):
    c = r1.cells[c_idx]
    c_bdr = c._tc.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders')
    c_xml = etree.tostring(c_bdr).decode('utf-8')
    if c_idx > 0:
        assert 'w:left w:val="double" w:sz="4"' in c_xml, f'Error: Cell {c_idx} left vertical must be double sz=4'
    if c_idx < len(r1.cells) - 1:
        assert 'w:right w:val="double" w:sz="4"' in c_xml, f'Error: Cell {c_idx} right vertical must be double sz=4'

print('Test 3 (Vertical lines all double sz=4): PASS')

# Check signer title in table 2 (bottom block)
t2 = doc.tables[2]
c_sign = t2.rows[0].cells[1]
p_role = c_sign.paragraphs[1]
print('Signer role text:', repr(p_role.text), 'bold:', p_role.runs[0].font.bold)
assert p_role.runs[0].font.bold is False, 'Error: Signer title CHANH VAN PHONG must NOT be bold'
print('Test 4 (Signer title not bold): PASS')

# Check HTML template
with open(r'D:\PM\LichCongTacTuan\templates\index.html', 'r', encoding='utf-8') as f:
    html_content = f.read()

assert '📄 Nạp file Word tuần cũ' not in html_content, 'Error: Emoji document icon still present in button'
assert 'history-search-input' not in html_content, 'Error: Search input still present in history modal'
assert 'file-import-json' not in html_content, 'Error: JSON import button still present in history modal'
print('Test 5 (UI button and history modal cleaned): PASS')

# 3. Convert to PDF and PNG for visual inspection
out_pdf = r'D:\PM\LichCongTacTuan\data\test_all_reqs.pdf'
out_png = r'D:\PM\LichCongTacTuan\data\test_all_reqs.png'

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
print('Test 6 (Visual PNG generated): PASS ->', out_png)
