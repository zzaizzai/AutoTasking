import glob
import pandas as pd
import os
import win32com.client as win32


class Muni:

    def __init__(self, target):
        dir_name = 'Muni'

        self.DesktopPath = os.path.expanduser('~/Desktop')
        self.target = target

        self.dir_data = self.DesktopPath + rf'\{target} Data'
        self.file_path_xls = self.DesktopPath + \
            fr'\{target} Data\{dir_name} {target}*.xls'
        self.file_xls = ''
        self.file_xlsx = ''

    def FindFile(self):
        print()
        print('Muni....')

        print('finding files')
        file_list = glob.glob(self.file_path_xls)
        print(file_list)

        if len(file_list) > 0:
            print('xls file exist')
            self.file_xls = file_list[0]
        else:
            print('no file')
            return

    def MakeXlsmFile(self):

        if self.file_xls == '':
            print('no xls file')
            return
        self.file_xlsx = self.file_xls + 'x'
        is_file = os.path.isfile(self.file_xlsx)

        if is_file:
            print('xlsx file exist')
            os.remove(self.file_xlsx)
            print('removed already existing xlsx file')
        else:   
            pass

        print('making xlsx file')
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(self.file_xls)

        wb.SaveAs(self.file_xlsx, FileFormat=51)
        wb.Close()
        excel.Application.Quit()

    def ReadFile(self):
        print('read file')
        df = pd.read_excel(self.file_xlsx, header=None)
        print(df)

        target_num = 0
        init_row = 0
        for i, value in enumerate(df[0]):
            print(f'{i}: {value}')
            if value == '特性値：':
                target_num = i - 1
                init_row = i + 4
                break
        print(f'number of target: {target_num}')
        print(f'samples start row: {init_row}')

        df_muni = df.loc[[init_row]]
        for i in range(0, target_num):
            df_muni = df_muni.append(df.loc[[init_row + 2 + 2*i]])
        print(df_muni)

        #  change row and col
        print('tanslated the col and row')
        df_muni = df_muni.transpose()
        df_muni = df_muni.loc[[2,3,4]]

        print(df_muni)

        df_muni.insert(0, 2,['M', 'M', 'min'])
        df_muni.insert(0, 1,['M1', 'Vm', 'ST'])
        df_muni.insert(0, 0,['Muni', 'Muni', 'Muni'])
        df_muni.reset_index(inplace= True, drop= True)
        df_muni = df_muni.T.reset_index(drop=True).T
        print(df_muni)

        self.WriteData(df_muni)
    
    def WriteData(self, df_input):
        print('writing data....')
        print(df_input)

        file_data = self.dir_data + fr'\{self.target} Data.xlsx'

        is_file = os.path.isfile(file_data)
        if is_file:
            pass
        else:   
            print(file_data)
            print('no data file')
            return
        
        df = pd.read_excel(file_data, header=None)
        df.drop([0, 1], axis = 1, inplace = True)
        df = df.T.reset_index(drop=True).T

        df = df.append(df_input, ignore_index = True)
        print(df)

        df.to_excel(file_data, index=True, header=True, startcol=1)


muni = Muni('CBA001')
muni.FindFile()
muni.MakeXlsmFile()
muni.ReadFile()
