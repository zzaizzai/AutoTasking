import os
import glob
import win32com.client as win32
import pandas as pd



class Muuni:

    def __init__(self, target):
        self.exp_name = 'Muni'

        self.DesktopPath = os.path.expanduser('~/Desktop')

        self.file_dir = self.DesktopPath  + rf'\{target} Data'
        self.target = target
        self.path_xls = self.file_dir + rf'\{self.exp_name} {target}*.xls'
        self.file_xls = ''
        self.file_xlsx = ''

    def FindFile(self):
        print('find files...')
        print(self.path_xls)

        file_list = glob.glob(self.path_xls)
        print(file_list)

        if len(file_list) > 0 :
            print(f'found {len(file_list)} {self.exp_name} file(s) ')

            for file in file_list:
                self.file_xls = file
                self.MAkeXlsmFile()
        else:
            print(f'No {self.exp_name}')
            return

    def MAkeXlsmFile(self):
        print('make xlsm file...')

        self.file_xlsx = self.file_xls + 'x'
        is_file = os.path.isfile(self.file_xlsx)

        if is_file:
            print('xlsx file exist')
            os.remove(self.file_xlsx)
        else:
            pass

        print('make new xlsx file...')
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(self.file_xls)
        wb.SaveAs(self.file_xlsx, FileFormat=51)
        wb.Close()
        excel.Application.Quit()

        print('make new xlsx file done')

        self.ReadFile()

    def ReadFile(self):
        print('read file...')
        print(self.file_xlsx)

        df = pd.read_excel(self.file_xlsx, header=None)
        print(df)
        for i, value in enumerate(df[0]):
            if value == '特性値：':
                num_target = i - 1
                row_init = i + 4
        
        print(f'number of target: {num_target}')
        print(f'samples start row: {row_init}')

        df_input = df.loc[[row_init]]

        print(df_input)

        

    def WriteData(self):
        print('writing data...')

        file_data = self.file_dir + fr'\{self.target} Data.xlsx'

        is_file = os.path.isfile(file_data)

        if is_file:
            pass
        else:
            print('no data file')
            # or you can make a data file
            return
        
        df = pd.read_excel(file_data, index_col=0)
        print(df)

muni = Muuni('CBA001')
muni.FindFile()