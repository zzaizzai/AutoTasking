import glob
from operator import index
import Service
import pandas as pd


class OilTension:

    def __init__(self, target:str):
        self.exp_name = '耐油引張り '

        self.file_dir = Service.desktop  + rf'\{target} Data'
        self.target = target

        self.file_path = self.file_dir + rf'\{self.exp_name}*{target}*.xls*'

        self.file_now = ''

    def FindFile(self):
        print('Find file')
        print(self.file_path)
        
        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)

        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s) ')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name} ')
            return
    def ReadFile(self):
        print('ReadFile')
        print('read file...')

        sheet_list = pd.ExcelFile(self.file_now).sheet_names
        print(sheet_list)

        if '設定シート' in  sheet_list:
            sheet_list.remove('設定シート')
        else:
            pass

        print(sheet_list)

        df_all = pd.DataFrame()

        for sheet in sheet_list:
            df_all = pd.concat([df_all, self.ReadDataSheet(sheet)])
        

        print('all df input data')

    def ReadDataSheet(self, sheet: str):
        print('read data shett')
        df = pd.read_excel(self.file_now, sheet_name=sheet, header=9, index_col=None)
        print(df)
        print(df.iloc[:,[1,9,10,11]])

        df_data = df.iloc[:,[1,9,10,11]]
        df_data = df_data.query("EB == EB")

        print(df_data)
def DoIt(target:str):
    oiru = OilTension(target)
    oiru.FindFile()


if __name__ == '__main__':
    target = input('target: ')
    print('Oil Tension')
    DoIt(target)
