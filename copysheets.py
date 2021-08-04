import openpyxl

path = r'WB.xlsx'
path2 = r'WB2.xlsx'

wb = openpyxl.load_workbook(path)
ws1 = wb.worksheets[0]

wb2 = openpyxl.load_workbook(path2)
ws2 = wb2.worksheets[0]

print(type(ws2['C2'].value))

for rows in ws1:
    for cells in rows:
        print(cells.value)
        celltype = type(cells.value) == type(ws2['C2'].value)
        print(celltype)
        if type(cells.value) == type(ws2['C2'].value):
            ws2[cells.coordinate].value = 'this is test for conditional copy'
        else:
             ws2[cells.coordinate].value = cells.value

print(ws2['C2'].value)
wb2.save(path2)
