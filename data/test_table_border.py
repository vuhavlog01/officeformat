import docx
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import os

def set_cell_border(cell, **kwargs):
    """
    kwargs can be top, bottom, left, right.
    value is dict(val='dotted'|'single'|'double'|'thinThickLargeGap', sz='4'|'12'|'24', color='auto'|'000000')
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

doc = docx.Document()
table = doc.add_table(rows=5, cols=4)
# Header
for c in table.rows[0].cells:
    set_cell_border(c, top={'val': 'thinThickLargeGap', 'sz': '24'}, bottom={'val': 'double', 'sz': '12'})

# Row 1 (Session 1, Item 1) -> dotted bottom
for c_i in [1, 2, 3]:
    set_cell_border(table.rows[1].cells[c_i], bottom={'val': 'dotted', 'sz': '4'})
    set_cell_border(table.rows[2].cells[c_i], top={'val': 'dotted', 'sz': '4'})

# Row 2 (Session 1, Item 2 - last) -> single bottom
for c_i in [1, 2, 3]:
    set_cell_border(table.rows[2].cells[c_i], bottom={'val': 'single', 'sz': '4'})

out_docx = os.path.abspath(r'D:\PM\LichCongTacTuan\data\test_custom_border.docx')
doc.save(out_docx)

# Convert to PDF
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
d = word.Documents.Open(out_docx)
out_pdf = os.path.abspath(r'D:\PM\LichCongTacTuan\data\test_custom_border.pdf')
d.SaveAs2(out_pdf, FileFormat=17)
d.Close()
word.Quit()
print("Success! Created test_custom_border.docx and pdf")
