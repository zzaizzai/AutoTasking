import pandas as pd
import Service
import glob
import os

class Compression:
    def __init__(self, target):
        self.target = target
        self.exp_name = '圧縮永久歪み_自動集積'
        self.file_path = Service.data_dir(target) + rf'\{self.exp_name}*{target}*.xls*'
        self.file_now = ''

    def FindFile(self):
        print('find file')
        print(self.file_path)
        
        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list,key=len)
        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s)')

            for file in file_list:
                self.file_now = file
                self.ReadFile()

        else:
            print(f'No {self.exp_name}')
            return

    def ReadFile(self):
        print('read File')
        print(self.file_now)

        sheet_list = pd.ExcelFile(self.file_now).sheet_names
        print(sheet_list)


        df_all = pd.DataFrame()

        for sheet in sheet_list:
            print(sheet)
            df_all = pd.concat([df_all, self.ReadDataSheet(sheet)])

        # print('all input data')

        # print(df_all)

        self.WriteData(df_all)

    def ReadDataSheet(self, sheet: str):
        print(sheet)
        df = pd.read_excel(self.file_now, sheet_name=sheet, header=9)
        print(df)
        df = df.iloc[:,[1,7]]

        title = ['配合番号','歪率']
        df.columns = title
        df['配合番号'] = df['配合番号'].ffill()

        # print(df)

        target_list = df['配合番号'].to_list()

        target_list_set = set(target_list)
        target_list = list(target_list_set)

        print(target_list)

        # find mean data

        mean_data_index = []
        for i in range(len(target_list)):
            mean_index = 3 + 4*i
            print(mean_index, df['配合番号'][mean_index])
            mean_data_index.append(mean_index)

        df = df.loc[mean_data_index]

        # rounding to int
        df = df.round(0)

        df =  df.transpose()
        df.columns = df.loc['配合番号']
        df.drop(index=['配合番号'], inplace=True)
        # print(df)

        unit = ['%']
        # method = ['compression']
        method = [Service.file_name_without_target(self.file_now, self.target)]
        condition = [sheet]
        type_list = ['distortion']

        df.insert(0, 'unit', unit )
        df.insert(0, 'type', type_list )
        df.insert(0, 'condition', condition )
        df.insert(0, 'method', method )
        df.reset_index(inplace=True, drop=True)

        print(df)

        return df

    def WriteData(self, df_input):
        print('merging data...')

        file_data = Service.data_dir(self.target)  + fr'\{self.target} Data.xlsx'
        is_file = os.path.isfile(file_data)

        if is_file:
            pass
        else:
            print('no data file')
            return
            
        Service.save_to_data_excel(file_data, df_input)

def DoIt(target: str):
    zumizumi = Compression(target)
    try:
        zumizumi.FindFile()
    except Exception as e:
        print(e)



if __name__ == '__main__':
    target = input('target: ')
    DoIt(target)
    