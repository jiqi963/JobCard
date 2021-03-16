# import library
from docx import Document
from docx.enum.section import WD_ORIENTATION
from docx.shared import Inches
from docx.shared import Pt
from docx.oxml.ns import qn

document = Document()

# set up page to landscape
section = document.sections[0]
section.orientation = WD_ORIENTATION.LANDSCAPE
page_h, page_w = section.page_width, section.page_height
section.page_width = page_w
section.page_height = page_h

# set up page margins
section.left_margin, section.right_margin = Inches(0.2), Inches(0.2)
section.top_margin, section.bottom_margin = Inches(0.2), Inches(0.2)

# section.sectPr.xpath('/w:cols')[0].set(qn('w:num'), '2')

table = document.add_table(rows=12, cols=2, style='Table Grid')

table.cell(1, 0).merge(table.cell(2, 0))
table.cell(3, 0).merge(table.cell(5, 0))
table.cell(7, 0).merge(table.cell(8, 0))
table.cell(9, 0).merge(table.cell(11, 0))

table.cell(0, 1).merge(table.cell(2, 1))
table.cell(3, 1).merge(table.cell(5, 1))
table.cell(6, 1).merge(table.cell(8, 1))

paragraph = document.add_paragraph()
paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = Pt(18)

table = document.add_table(rows=12, cols=2, style='Table Grid')
table.cell(1, 0).merge(table.cell(2, 0))
table.cell(3, 0).merge(table.cell(5, 0))
table.cell(7, 0).merge(table.cell(8, 0))
table.cell(9, 0).merge(table.cell(11, 0))

table.cell(0, 1).merge(table.cell(2, 1))
table.cell(3, 1).merge(table.cell(5, 1))
table.cell(6, 1).merge(table.cell(8, 1))

document.save('test.docx')