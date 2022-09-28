import os
import glob
import pandas as pd
class HeatResist:

    def __init__(self, target: str):
        self.exp_name = '熱老化_自動集積 '

        self.DesktopPath = os.path.expanduser('~/Desktop')

        self.file_dir = self.DesktopPath  + rf'\{target} Data'
        self.target = target

        self.file_path = self.file_dir + rf'\{self.exp_name}*{target}*.xls*'

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

        if '設定シート' in sheet_list:
            sheet_list.remove('設定シート')
        else:
            pass

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
        df = pd.read_excel(self.file_now, sheet_name=sheet, header=None, index_col=1)
        df = df.transpose()
        print(df.index)


        print(df)



        ## title
        print(df)
        print(df.iloc[:,[2,3]])

        print(df.loc[[4]].values.tolist()[0])
        col_index = []
        for i, value in enumerate(df.loc[[4]].values.tolist()[0]):
            print(i, value)
            if not 'nan' in str(value):
                print('not nan')
                col_index.append(i)

        print(col_index)
        print(df.iloc[:,col_index])

        df = df.iloc[:,col_index]

        title_index =  df.columns.values.tolist()
        for i, value in enumerate(title_index):
            if 'nan' in str(value):
                print('nan',i)
                title_index[i] = title_index[i-1]
        print(title_index)
        df.columns = title_index
        print(df)



        
        mean_col_index = [0]
        row_mean_str = df.loc[[3]].values.tolist()[0]
        print(row_mean_str)
        for i, value in enumerate(row_mean_str):
            print(value)
            if '中央値' in str(value):
                print('mean str', i)
                mean_col_index.append(i)

        print(mean_col_index)

        df_input = df.iloc[:,mean_col_index]
        df_input = df_input.dropna()
        print(df_input)

        print(df_input.query("配合番号 in ['引っ張り','伸び']"))

        return df


def DoIt(target:str):
    heat = HeatResist(target)
    heat.FindFile()

if __name__ == '__main__':
    target = 'CBA001'
    # target = input('target: ')
    
    DoIt(target)
