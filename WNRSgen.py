from docx import Document
from docx.shared import Inches
from docx.shared import RGBColor
from docx.shared import Pt
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH

def fillincells():
    for k in range(len(recfixed)):
        row = table.add_row().cells
        for cells in table.rows[-1].cells:
            cells.text = lines[k]
            run = cells.paragraphs[0].runs[0]  # formatting occurs at run level, I'll call it as real-time
            run.font.bold = True
            run.font.name = 'Calibri'
            run.font.size = Pt(24)
            run.font.color.rgb = RGBColor(0x7E, 0x35, 0x17)

lines = []
with open('wnrslines.txt','r') as f:
    lines = f.readlines()
print(lines)

count = 0
for line in lines:
    count += 1
print("number of lines: "+str(count))

document = Document()

recfixed = tuple(lines)
print(recfixed)

for i in range(len(recfixed)):
    print(i)
    i += 1

table = document.add_table(rows=0, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = 'Table Grid'
# paragraph = document.add_paragraph()
# paragraph_format = paragraph.paragraph_format
# paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

fillincells()
for row in table.rows:
    row.height = Cm(5.6)

table.cell(0, 0).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
table.cell(0, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
table.cell(1, 0).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
table.cell(1, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER #horizontal allignment

table.cell(1, 1).vertical_alignment = WD_TABLE_ALIGNMENT.CENTER #vertical allignment

document.save('WNRSprint.docx')