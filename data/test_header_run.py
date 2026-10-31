import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client, os, fitz

doc = docx.Document()
section = doc.sections[0]
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.5)
section.left_margin = Cm(2.0)
section.right_margin = Cm(1.5)

table = doc.add_table(rows=1, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False
table.columns[0].width = Cm(7.5)
table.columns[1].width = Cm(10.0)

cell_left = table.cell(0, 0)
cell_right = table.cell(0, 1)

# Left
p1 = cell_left.paragraphs[0]
p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p1.paragraph_format.space_before = Pt(0)
p1.paragraph_format.space_after = Pt(2)
r1 = p1.add_run('ĐẢNG BỘ THÀNH PHỐ ĐÀ NẴNG')
r1.font.name = 'Times New Roman'
r1.font.size = Pt(12)

p2 = cell_left.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_before = Pt(0)
p2.paragraph_format.space_after = Pt(2)
r2 = p2.add_run('ĐẢNG UỶ XÃ QUẾ SƠN')
r2.font.name = 'Times New Roman'
r2.font.size = Pt(12.5)
r2.font.bold = True

p_star = cell_left.add_paragraph()
p_star.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_star.paragraph_format.space_before = Pt(0)
p_star.paragraph_format.space_after = Pt(4)
r_star = p_star.add_run('*')
r_star.font.name = 'Times New Roman'
r_star.font.size = Pt(12)
r_star.font.bold = True

# Right
p_dang = cell_right.paragraphs[0]
p_dang.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_dang.paragraph_format.space_before = Pt(0)
p_dang.paragraph_format.space_after = Pt(2)
r3 = p_dang.add_run('ĐẢNG CỘNG SẢN VIỆT NAM')
r3.font.name = 'Times New Roman'
r3.font.size = Pt(14)
r3.font.bold = True

# Line under ĐẢNG CỘNG SẢN VIỆT NAM
p_rule = cell_right.add_paragraph()
p_rule.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_rule.paragraph_format.space_before = Pt(0)
p_rule.paragraph_format.space_after = Pt(3)
p_rule.paragraph_format.line_spacing = Pt(1)
p_rule.paragraph_format.left_indent = Cm(1.9)
p_rule.paragraph_format.right_indent = Cm(1.9)
pPr = p_rule._p.get_or_add_pPr()
pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="1" w:color="000000"/></w:pBdr>')
pPr.append(pBdr)

# Date
p_date = cell_right.add_paragraph()
p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_date.paragraph_format.space_before = Pt(2)
p_date.paragraph_format.space_after = Pt(6)
r4 = p_date.add_run('Quế Sơn, ngày 1 tháng 11 năm 2026')
r4.font.name = 'Times New Roman'
r4.font.size = Pt(13)
r4.font.italic = True

out_docx = r'D:\PM\LichCongTacTuan\data\test_header.docx'
out_pdf = r'D:\PM\LichCongTacTuan\data\test_header.pdf'
out_png = r'D:\PM\LichCongTacTuan\data\test_header.png'
doc.save(out_docx)

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
print('Success! Check:', out_png)
