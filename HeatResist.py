import os
import glob
import pandas as pd
class HeatResist:

    def __init__(self, target: str):
        self.exp_name = '熱老化_自動集積 '

        self.DesktopPath = os.path.expanduser('~/Desktop')

        self.file_dir = self.DesktopPath  + rf'\{target} Data'
        self.target = target

        self.file_path = self.file_dir + rf'\{self.exp_name}*{target}*.xlsm'

        self.file_now = ''
    
    def FindFile(self):
        print('find files...')
        print(self.file_path)

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)
        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s)')

            for file in file_list:
                self.file_now = file
                self.ReadFile()


        else: print(f'No {self.exp_name}')
        return

    def ReadFile(self):
        print('read file...')
        print(self.file_now)
        sheet_list = pd.ExcelFile(self.file_now).sheet_names
        print(sheet_list)
        sheet_list.remove('設定シート')
        print(sheet_list)


        df_all = pd.DataFrame()

        for sheet in sheet_list:
            df_all = pd.concat([df_all, self.ReadDataSheet(sheet)])
        

        # print(df_all)

    def ReadDataSheet(self, sheet: str):

        def handleData():
            # make df simple row and col
            print()

            
        print(sheet)
        df = pd.read_excel(self.file_now, sheet_name=sheet, header=10, index_col=1)
        df = df.transpose()
        print(df.index)


        print(df)

        return df


def DoIt(target:str):
    heat = HeatResist(target)
    heat.FindFile()

if __name__ == '__main__':
    target = 'FJX001'
    # target = input('target: ')
    
    DoIt(target)
