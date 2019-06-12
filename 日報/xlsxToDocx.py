import docx
import xlrd
import pprint
import xlwt
from docx.shared import RGBColor
from docx.shared import Pt

sheetName = input("ファイル名(xxx)を入力(xxx.xlsx形式)>>> ")
doc = docx.Document()
wb = xlrd.open_workbook(sheetName +".xlsx")
sheet = wb.sheet_by_index(0)
################################################
row = sheet.row_values(0)
cols = []
for i, e in enumerate(row):
    col = sheet.col_values(i)
    cols.append(col)

for i, e in enumerate(row):
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
doc.save('./{}.docx'.format(sheetName))
print("fin")
