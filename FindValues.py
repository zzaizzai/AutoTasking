import openpyxl
import glob


target = 'CBA001'

filePath = fr'C:\Users\junsa\Desktop\CBA001\Oil {target}.*'
list = glob.glob(filePath, recursive=True)
wb = openpyxl.load_workbook(list[0])
sheet = wb.worksheets[0]
print(sheet)

targetCell: str = 'FS'

def ValuePosition(cellValue):
    for row in sheet.rows:
        for cell in row:
            if cell.value != None and cell.value == cellValue:
                    # print('it is target')
                    # print(f'{cell.value}: {sheet.cell(row= cell.row , column= cell.column + 1).value}')
                    # print(f"{cell.value}: {cell.coordinate}")
                    return cell


print(f'x: {ValuePosition(targetCell).row}, y : {ValuePosition(targetCell).column}')

wb.close()