import docx
import xlrd
import pprint
import xlwt
from docx.shared import RGBColor
from docx.shared import Pt
#####__init__###################################
sheetName = input("ファイル名(xxx)を入力(xxx.xlsx形式)>>> ")
doc = docx.Document()
wb = xlrd.open_workbook(sheetName +".xlsx")
sheet = wb.sheet_by_index(0)
#####
section = doc.sections[0]
header = section.header
paragraph = header.paragraphs[0]
# paragraph.size = Pt(16)
paragraph.text = sheetName
################################################
row = sheet.row_values(0)
cols = []
for i, e in enumerate(row):
    col = sheet.col_values(i)
    cols.append(col)

for i, e in enumerate(row):
    if i == 0:
        continue
    else:
        for ii, ee in enumerate(cols[i]):
            # print("{}>>> {}".format(i, cols[i][ii]))
            if ii == 0:
                doc.add_heading(cols[i][ii], 0)
            else:
                if str(cols[i][ii]) == "":
                    continue
                else:
                    # print(cols[len(cols)-1][ii].color.type)
                    # doc.add_paragraph(str(cols[len(cols)-1][ii]) +":")
                    #####
                    run = doc.add_paragraph().add_run(str(cols[len(cols)-1][ii]) +"  :")
                    font = run.font
                    font.color.rgb = RGBColor(47, 145, 79)
                    font.size = Pt(8)
                    #####
                    doc.add_paragraph("   " +str(cols[i][ii]))
################################################
paragraph.text = sheetName +" #出席人数:{}人".format(len(cols[len(cols)-1]))
doc.save('./{}.docx'.format(sheetName))
print("fin")
