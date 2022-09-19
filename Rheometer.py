import openpyxl
import glob
import win32com.client as win32
import os
import pandas as pd
import numpy as np


class Rheometer:
    def __init__(self, target, DesktopPath:str):
        self.target = target
        self.folderPath = DesktopPath + fr'\{target} Data'
        self.filePath = DesktopPath + fr'\{target} Data\レオメータ*{target}*.xls'
        self.number_of_target: int = 0

    # translate from xls to xlsx
    def MakeXlsmFile(self, file: str):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(file)
        xlsxFile = file + 'x'
        wb.SaveAs(xlsxFile, FileFormat=51)
        wb.Close()  # FileFormat = 56 is for .xls extension
        excel.Application.Quit()

    def CreateGraph(self, xlsxFilePath: str):
        wb = openpyxl.load_workbook(xlsxFilePath)
        # wb = openpyxl.load_workbook(xlsxFile , keep_vba=True)
        # ws = wb['Sheet1']
        ws = wb.worksheets[0]


        standard_cell = ws.cell(row=1, column=1)
        # Find 'Time' cell
        for row in ws.rows:
            for cell in row:
                if cell.value == 'Time(NO.1)':
                    standard_cell = cell
                    break
        
        self.number_of_target = int(ws.cell(row=standard_cell.row - 2, column=standard_cell.column).value)
        print(f'number of target: {self.number_of_target}')

        print(f'standard postion cell: {standard_cell}')
        # 1/4 data
        # print('1/4 data....')
        # from_row = standard_cell.row + 20
        # for i in range(200):
        #     print(f'deleting{i}')
        #     ws.delete_rows( from_row + i*3 + 3)
        #     ws.delete_rows( from_row + i*3 + 2)
        #     ws.delete_rows( from_row + i*3 + 1)
        #     ws.delete_rows( from_row + i*3)

        # create a graph
        chart = openpyxl.chart.ScatterChart('marker')

        # data titles
        chart.title = 'Rheometer'
        chart.x_axis.title = 'Time'
        chart.y_axis.title = 'torque'

        # add series
        print('adding serials')
        for i in range(self.number_of_target):
            serial_name = ws.cell(row=21, column=2 + 4*i).value
            if serial_name != None:
                print(serial_name)
                data = openpyxl.chart.Reference(
                    ws, min_col=2 + i*4, min_row=standard_cell.row, max_row=1000)
                times = openpyxl.chart.Reference(
                    ws, min_col=1, min_row=standard_cell.row + 1, max_row=1000)

                series = openpyxl.chart.Series(
                    data, times, title_from_data=True)
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

        print('graph save done!!')

        self.CopyValues(xlsxFilePath)

    def Rheometer(self):
        print('Searching Rheometer Data....')
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

    def CopyValues(self, xlsxFilePath):
        print(xlsxFilePath)
        print('copying values...')
        wb = openpyxl.load_workbook(xlsxFilePath)
        ws = wb.worksheets[0]

        # make values name
        init_row = 1
        init_col = 2

        xlsxFilePath_copy = os.path.dirname(
            xlsxFilePath) + f'\{self.target} Data.xlsx'
        wb_copy = openpyxl.load_workbook(xlsxFilePath_copy)
        ws_copy = wb_copy.worksheets[0]

        # targets
        for i in range(self.number_of_target):
            # print(f'{self.target}({i}) making')
            ws_copy.cell(row=1, column=5 + i,  value=f'{self.target}({i+1})')

        ws_copy.cell(row=init_row, column=init_col, value='Experiemnts')
        ws_copy.cell(row=init_row, column=init_col + 1, value='Method')
        ws_copy.cell(row=init_row, column=init_col + 2, value='Unit')

        method = ['MH', 'ML', 't10', 't50', 't90', 'CR']
        for i in range(len(method)):
             ws_copy.cell(row=init_row + 1 + i, column=init_col + 1, value=method[i])
        # ws_copy.cell(row=init_row + 1, column=init_col + 1, value='MH')
        # ws_copy.cell(row=init_row + 2, column=init_col + 1, value='ML')
        # ws_copy.cell(row=init_row + 3, column=init_col + 1, value='t10')
        unit = ['kgf・cm', 'kgf・cm', 'min', 'min', 'min']
        ws_copy.cell(row=init_row + 1, column=init_col + 2, value='%')
        ws_copy.cell(row=init_row + 2, column=init_col + 2, value='%')
        ws_copy.cell(row=init_row + 3, column=init_col + 2, value='%')

        for i in range(3):
            ws_copy.cell(row=init_row + 1 + i, column=init_col, value='Rheometer')

        cell_init = ws_copy.cell(row=1, column=1)
        # Find top of the value postion
        breaker = False
        for row in ws.rows:
            for cell in row:
                if cell.value == 'MH':
                    print(cell)
                    cell_init = cell
                    breaker = True
                    break
            if breaker == True:
                break

        print(f'standard cell: {cell_init}')

        def values_original(row, col):
            ws.cell(row=row, column=col).value
            return ws.cell(row=row, column=col).value

        for i in range(5):
            ws_copy.cell(
                row=2, column=5 + i, value=values_original(cell_init.row + 3 + i*2, cell_init.column))
            ws_copy.cell(
                row=3, column=5 + i, value=values_original(cell_init.row + 3 + i*2, cell_init.column + 1))
            ws_copy.cell(
                row=4, column=5 + i, value=values_original(cell_init.row + 3 + i*2, cell_init.column + 3))

        wb_copy.save(xlsxFilePath_copy)
        print('saved copy values of rheometer')

    # def removeNanCol(self):
    #     wb = openpyxl.Workbook()
    #     ws = wb.active

    #     # delete nan value col
    #     excelPath = self.folderPath + fr'\{self.target} Data.xlsx'
    #     df = pd.read_excel(excelPath, header=None, index_col=0)
    #     df.reset_index(inplace= True, drop= True)
    #     # df.loc[len(df)] = data_input
    #     print(df)
    #     print(len(df.columns))

    #     for i in range(4, len(df.columns) + 1):
    #         print(df[i])
    #         # np.isnan(df[i][2])
    #         if pd.isna(df[i][2]):
    #         # if np.isnan(df[i][2]):
    #             del df[i]

    #     df.to_excel(excelPath, index=False, header=False, startcol=1)
    #     # wb.save(self.folderPath + fr'\{self.target} Data.xlsx')


def Rheo(target: str, DesktopPath:str):
    
    file_list =  glob.glob(DesktopPath + fr'\{target} Data\レオメータ*{target}*.xls')
    if len(file_list) > 0:
        print()
        print()
        print('rheometer file exists')
        print(file_list)
        rheo = Rheometer(target, DesktopPath)
        rheo.Rheometer()
        # rheo.removeNanCol()
    else:
        print('no rheometer Data')



if __name__ == 'Rheometer.py':
    Rheo()
