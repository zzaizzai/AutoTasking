import openpyxl


target = 'FJX001'
# filePath = fr'C:\Users\junsa\Desktop\CBA001\Oil\{target}.xlsm'
filePath = r'C:\Users\1010020990\Desktop\FJX001 data\レオメータ_東洋3号4号_自動集積\FJX001-005.xls'
import win32com.client as win32
fname = filePath
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)
wb.SaveAs(fname+"x", FileFormat = 51)
wb.Close() #FileFormat = 56 is for .xls extension
excel.Application.Quit()

# translate xls to xlsx
# pyexcel.save_book_as(file_name=filePath, dest_file_name=r'C:\Users\1010020990\Desktop\FJX001 data\レオメータ_東洋3号4号_自動集積\FJX001-005 done.xlsx')

# wb = xlrd.open_workbook(filePath)
wb = openpyxl.load_workbook(r'C:\Users\1010020990\Desktop\FJX001 data\レオメータ_東洋3号4号_自動集積\FJX001-005.xlsx' , keep_vba=True)
ws = wb.worksheets[0]


# 1/4 data
for i in range(335):
    ws.delete_rows( 25 + i*3)

chart = openpyxl.chart.LineChart()
# chart.style = 26
# colors 10, 18, 26

chart.title = 'Rheometer'
# data titles
chart.x_axis.title = 'Time'
chart.y_axis.title = 'torque'


# data scail
chart.y_axis.scaling.min = 0
chart.y_axis.scaling.max = 35

chart.x_axis.scaling.min = 10
chart.x_axis.scaling.max = 35




# input data

# data including data title
for i in range(5):
    data = openpyxl.chart.Reference(ws, min_col =2 + i*4, min_row =21, max_row = 1024)
    chart.add_data(data, from_rows = False, titles_from_data = True)

# categories
times = openpyxl.chart.Reference(ws, min_col = 1, min_row =22 , max_row = 1024)
chart.set_categories(times)

print(len(chart.series))
for ss in chart.series:
    ss.marker.symbol = "circle"
    ss.graphicalProperties.line.noFill=True 

ws.add_chart(chart, 'B8')
wb.save(r'C:\Users\1010020990\Desktop\FJX001 data\レオメータ_東洋3号4号_自動集積\FJX001-005.xlsm')
print('done')
