import os
import glob
import win32com.client as win32
import pandas as pd



class Muuni:

    def __init__(self, target):
        self.exp_name = 'ムーニー_ロータ_自動集積'

        self.DesktopPath = os.path.expanduser('~/Desktop')

        self.file_dir = self.DesktopPath  + rf'\{target} Data'
        self.target = target
        self.path_xlsx = self.file_dir + rf'\{self.exp_name} {target}*.xlsx'
        self.file_xlsx = ''
        self.index_name = ''

    def FindFile(self):
        print('find files...')
        print(self.path_xlsx)

        file_list = glob.glob(self.path_xlsx)
        file_list = sorted(file_list, key=len)
        print(file_list)

        if len(file_list) > 0 :
            print(f'found {len(file_list)} {self.exp_name} file(s) ')

            for file in file_list:
                self.file_xlsx = file
                self.MAkeXlsmFile()
        else:
            print(f'No {self.exp_name}')
            return

    def MAkeXlsmFile(self):
        # print('make xlsm file...')

        # self.file_xlsx = self.file_xls + 'x'
        # is_file = os.path.isfile(self.file_xlsx)

        # if is_file:
        #     print('xlsx file exist')
        #     os.remove(self.file_xlsx)
        # else:
        #     pass

        # print('make new xlsx file...')
        # excel = win32.gencache.EnsureDispatch('Excel.Application')
        # wb = excel.Workbooks.Open(self.file_xls)
        # wb.SaveAs(self.file_xlsx, FileFormat=51)
        # wb.Close()
        # excel.Application.Quit()

        # print('make new xlsx file done')

        self.ReadFile()

    def ReadFile(self):
        print('read file...')
        print(self.file_xlsx)

        df = pd.read_excel(self.file_xlsx, header=None)
        print(df)
        for i, value in enumerate(df[0]):
            if value == '特性値：':
                num_target = df[0][i-1]
                row_init = i + 4
        
        print(f'number of target: {num_target}')
        print(f'samples start row: {row_init}')

        df_input = df.loc[[row_init]]
        for i in range(1, num_target):
            df_input = df_input.append(df.loc[[row_init + 2*i]])

        print('translate row and col')
        df_input = df_input.transpose()
        print(df_input)

        print('')
        df_input = df_input.loc[[2,3,4]]
        print(df_input)

        file_name = os.path.splitext(os.path.basename(self.file_xlsx))[0]
        print(file_name)
        unit = ['M', 'M', 'min' ]
        method = ['M1', 'Vm', 'T1']
        condition = ['none','none','none']
        name = [file_name, file_name, file_name]
        df_input.insert(0, 3, unit)
        df_input.insert(0, 2, method)
        df_input.insert(0, 1, condition)
        df_input.insert(0, 0, name)

        print(df_input)
        
        # reset title and index
        df_input.reset_index(inplace= True, drop= True)
        df_input = df_input.T.reset_index(drop=True).T

        print(df_input)

        self.WriteData(df_input)

    def WriteData(self, df_input) :
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

        df_merge = pd.concat([df, df_input])
        print(df_merge)

        df_merge.reset_index(inplace= True, drop= True)

        df_merge.to_excel(file_data, index=True, header=True, startcol=0)
        print(f'saved data file in {file_data}')

def DoIt(target: str):
    muni = Muuni(target)
    muni.FindFile()



if __name__ == '__main__':
    target = target = input('target number (ex: ABC001): ')
    DoIt(target)
