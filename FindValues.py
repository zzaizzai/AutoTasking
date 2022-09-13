import openpyxl



filePath = r'C:\Users\junsa\Desktop\CBA001\Delta CBA001.xlsx'
wb = openpyxl.load_workbook(filePath)
sheet = wb.worksheets[0]
print(sheet)

for row in sheet.rows:
    for cell in row:
        if cell.value != None:
            print(f"{cell} : {cell.value}")

wb.close()