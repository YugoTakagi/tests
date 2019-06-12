import docx
import xlrd
import pprint
import xlwt

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

    for ii, ee in enumerate(cols[i]):
        # print("{}>>> {}".format(i, cols[i][ii]))
        if ii == 0:
            doc.add_heading(cols[i][ii], 0)
        else:
            doc.add_paragraph(str(cols[i][ii]))
################################################
doc.save('./{}.docx'.format(sheetName))
