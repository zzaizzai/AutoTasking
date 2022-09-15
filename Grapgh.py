import matplotlib.pylab as plt
import openpyxl

target = 'CBA001'
filePath = fr'C:\Users\junsa\Desktop\CBA001\Oil\{target}.xlsm'

wb = openpyxl.load_workbook(filePath)
ws = wb['Sheet1']
for i in range(10):
    ws.delete_rows( 3 + i*3)
    # 1/4 data

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

chart.x_axis.scaling.min = 0
chart.x_axis.scaling.max = 35




# input data

# data including data title
for i in range(5):
    data = openpyxl.chart.Reference(ws, min_col =3 + i*2, min_row =1, max_row = 30)
    chart.add_data(data, from_rows = False, titles_from_data = True)

# categories
times = openpyxl.chart.Reference(ws, min_col = 2, min_row =2 , max_row = 31)
chart.set_categories(times)

print(len(chart.series))
for ss in chart.series:
    ss.marker.symbol = "circle"
    ss.graphicalProperties.line.noFill=True 

ws.add_chart(chart, 'B8')
wb.save(fr'C:\Users\junsa\Desktop\CBA001\Oil\{target} done.xlsx')
print('done')