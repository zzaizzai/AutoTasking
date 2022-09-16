import openpyxl
import glob
import win32com.client as win32
import os


class Rheometer:
    def __init__(self, target):
        self.target = target
        self.filePath = fr'C:\Users\junsa\Desktop\CBA001\Oil\{target}*.xls'

    # translate xls to xlsx
    def MakeXlsmFile(self, file: str):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(file)
        xlsxFile = file +  'x'
        wb.SaveAs(xlsxFile, FileFormat = 51)
        wb.Close() #FileFormat = 56 is for .xls extension
        excel.Application.Quit()

    def CreateGraph(self, xlsxFilePath: str):
        wb = openpyxl.load_workbook(xlsxFilePath)
        # wb = openpyxl.load_workbook(xlsxFile , keep_vba=True)
        ws = wb['Sheet1']

        # 1/4 data
        for i in range(500):
            if i%4 != 0:
                ws.delete_rows( 3 + i)

        # create a graph
        chart = openpyxl.chart.ScatterChart('marker')

        # data titles
        chart.title = 'Rheometer'
        chart.x_axis.title = 'Time'
        chart.y_axis.title = 'torque'

        #add series
        for i in range(5):
            data = openpyxl.chart.Reference(ws, min_col =3 + i*2, min_row =1, max_row = 550)
            times = openpyxl.chart.Reference(ws, min_col = 2, min_row =2 , max_row = 550)

            series = openpyxl.chart.Series(data, times, title_from_data=True)
            series.graphicalProperties.line.noFill = True
            series.marker.symbol = "circle"

            chart.series.append(series)

        print(f'number of series: {len(chart.series)}')
        
        # # data scail
        # chart.y_axis.scaling.min = 0
        # chart.y_axis.scaling.max = 35

        # chart.x_axis.scaling.min = 0
        # chart.x_axis.scaling.max = 35

        ws.add_chart(chart, 'B8')
        wb.save(xlsxFilePath)

        print('graph save done')


    def Rheometer(self):
        print('Rheometer Data system...')
        files = glob.glob(self.filePath)
        print(f'candidate files: {files}')
        for file in files:
            print(f'handling: {file}')
            # delete exist xlsx file
            if glob.glob(file + 'x'):
                os.remove(file + 'x')
                print('deleted xlsx file')

            # making new xlsx file
            print('making a xlsx file')
            self.MakeXlsmFile(file)
            self.CreateGraph(file + 'x')

rheo = Rheometer('CBA001')
rheo.Rheometer()