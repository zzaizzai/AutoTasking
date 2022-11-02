import Service
import glob
import pandas as pd


class Ozone:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target: str):
        self.exp_name = 'オゾン '

        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_data = Service.data_dir(
            target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''
    def StartProcess(self):
        print('Start Process')

        file_list = glob.glob(self.file_path)
        file_list  = sorted(file_list, key=len)
        print(file_list)

        if len(file_list) > 0 :
            print(f'found {len(file_list)} file(s)')
            pass
        else:
            print(f'No {self.exp_name}')
            return
        

        for file in file_list:
            self.file_now = file
            self.ReadData()
    
    def ReadData(self):
        print('reading files')
        print(self.file_now)
        sheet_list = pd.ExcelFile(self.file_now).sheet_names

        if '設定シート' in sheet_list:
            sheet_list.remove('設定シート')

        if 'きれつ表' in sheet_list:
            sheet_list.remove('きれつ表')


        df_all = pd.DataFrame()

        for sheet in sheet_list:
            try:
                df_all = pd.concat([df_all, self.ReadSheet(sheet)])
            except Exception as e:
                print(e)
    def ReadSheet(self, sheet):
        print('reading sheet', sheet)


        df_sheet = pd.read_excel(self.file_now, sheet_name=sheet, header=10, index_col=1)
        df_sheet = df_sheet.iloc[:, 3:]
        df_sheet = df_sheet.dropna(how='all', axis=1)
        
        target_list = df_sheet.index.to_list()
        for i, value in enumerate(target_list):
            print(i, value)
            if str(value) == 'nan' and i > 0:
                target_list[i] = target_list[i-1]
    
        print(target_list)
        df_sheet.index = target_list
        print(df_sheet)

        df_n1 =  df_sheet.loc[:, ~df_sheet.columns.duplicated(keep="frist")]
        df_n2 =  df_sheet.loc[:, ~df_sheet.columns.duplicated(keep="last")]

        print(df_n1)
        print(df_n2)


def ozozo(target:str, test_mode =False):
    print('ozoozo')
    zozoni = Ozone(target)
    zozoni.TestMode(test_mode)
    zozoni.StartProcess()


if __name__ == "__main__":
    target = input('target: ')
    ozozo(target, test_mode=True)

