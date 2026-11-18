import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_border_bottom(cell, val="dotted", sz="4", color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    border_xml = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tcBorders>')
    tcPr.append(border_xml)

def set_cell_border_top(cell, val="dotted", sz="4", color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    border_xml = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tcBorders>')
    tcPr.append(border_xml)

doc = docx.Document()
tbl = doc.add_table(rows=3, cols=3)
# Row 0 to 1 border dotted in col 1, 2
set_cell_border_bottom(tbl.cell(0, 1), "dotted")
set_cell_border_top(tbl.cell(1, 1), "dotted")
set_cell_border_bottom(tbl.cell(0, 2), "dotted")
set_cell_border_top(tbl.cell(1, 2), "dotted")

doc.save(r'D:\PM\LichCongTacTuan\data\test_dotted.docx')
print("Successfully created test_dotted.docx")
