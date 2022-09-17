import openpyxl
import glob
import win32com.client as win32
import os


class Rheometer:
    def __init__(self, target):
        self.target = target
        self.filePath = fr'C:\Users\junsa\Desktop\{target} Data\Rheometer {target}*.xls'

    # translate from xls to xlsx
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

        standard_cell = ws.cell(row=1, column=1)
        # Find 'Time' cell
        for row in ws.rows:
            for cell in row:
                if cell.value == 'Time':
                    standard_cell = cell
                    break

        print(f'standard postion cell: {standard_cell}')
        # 1/4 data
        print('1/4 data....')
        from_row = standard_cell.row + 10
        for i in range(300):
            ws.delete_rows( from_row + i*2 + 6)
            ws.delete_rows( from_row + i*2 + 4)
            ws.delete_rows( from_row + i*2 + 2)
            ws.delete_rows( from_row + i*2)
           
        # create a graph
        chart = openpyxl.chart.ScatterChart('marker')

        # data titles
        chart.title = 'Rheometer'
        chart.x_axis.title = 'Time'
        chart.y_axis.title = 'torque'

        #add series
        print('adding serials')
        for i in range(8):
            serial_name = ws.cell(row=21, column=2 + 4*i).value
            if serial_name != None:
                print(serial_name)
                data = openpyxl.chart.Reference(ws, min_col =2 + i*4, min_row = standard_cell.row, max_row = 550)
                times = openpyxl.chart.Reference(ws, min_col = 1, min_row = standard_cell.row + 1 , max_row = 550)

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

        xlsxFilePath_copy = os.path.dirname(xlsxFilePath) + f'\{self.target} Data.xlsx'
        wb_copy = openpyxl.load_workbook(xlsxFilePath_copy)
        ws_copy = wb_copy.worksheets[0]

        ws_copy.cell(row=2, column=2, value='MA')
        ws_copy.cell(row=3, column=2, value='MB')
        ws_copy.cell(row=4, column=2, value='MC')
        for i in range(3):
            ws_copy.cell(row=2 + i, column=1, value='Rheometer')



        cell_init = ws_copy.cell(row=1, column=1)
        # Find top of the value postion
        breaker = False
        for row in ws.rows:
            for cell in row:
                if cell.value == 'MA':
                    print(cell)
                    cell_init = cell
                    breaker = True
                    break
            if breaker == True:
                break
                
        print(f'standard cell: {cell_init}')

        def values_original(row, col):
            ws.cell(row= row, column= col).value
            return ws.cell(row= row, column= col).value
            

        for i in range(5):
            ws_copy.cell(row=2 , column= 4 + i, value= values_original(cell_init.row + 1 + i*2 , cell_init.column))
            ws_copy.cell(row=3 , column= 4 + i, value= values_original(cell_init.row + 1 + i*2, cell_init.column + 2))
            ws_copy.cell(row=4 , column= 4 + i, value= values_original(cell_init.row + 1 + i*2, cell_init.column + 4))

        wb_copy.save(xlsxFilePath_copy)   
        print('saved copy values of rheometer')            


def Rheo(target:str):
    rheo = Rheometer(target)
    rheo.Rheometer()    

if __name__ == 'Rheometer.py':
    Rheo()

