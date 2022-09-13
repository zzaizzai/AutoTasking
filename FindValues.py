import openpyxl



filePath = r'C:\Users\junsa\Desktop\CBA001\Delta CBA001.xlsx'
wb = openpyxl.load_workbook(filePath)
sheet = wb.worksheets[0]
print(sheet)

targetCell: str = 'aa30'

for row in sheet.rows:
    for cell in row:
        if cell.value != None:
            if cell.value == targetCell:
                print('it is target')
                print(f'{cell.value}: {sheet.cell(row= cell.row , column= cell.column + 1).value}')
            # print(f"{cell.coordinate}: {cell.value}")


wb.close()