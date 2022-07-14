from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT

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
for k in range(len(recfixed)):
    table.add_row()
    for cells in table.rows[-1].cells:
        cells.text = lines[k]

document.save('WNRSprint.docx')