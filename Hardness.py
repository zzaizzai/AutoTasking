from matplotlib.pyplot import title
import pandas as pd
import Service
import os
import glob

# file = r'C:\Users\1010020990\Desktop\FJX010 Data\硬度_自動集積 FJX010.xlsx'
# target = 'FJX010'

# df = pd.read_excel(file, header=7)
# print(df.iloc[:,13:16])
# taregt_2 = df.iloc[:,13].to_list()

# for i in range(len(taregt_2)):
#     if str(taregt_2[i]) != "nan":
#         taregt_2[i] = Service.target_number(i, target)
# print(taregt_2)
# df.iloc[:,13] =taregt_2


# df = df.iloc[:,13:16]
# print(df)

# titles = df.columns.to_list()
# print(titles)
# titles = ['配合番号', '０秒','3秒']
# df.columns = titles

# print(df)

# df.dropna(how='all', inplace=True)
# df.dropna(how='all',axis=1,  inplace=True)
# print(df)

# df = df.transpose()
# # df = df[:,1]
# titles_new = df.loc['配合番号'].to_list()
# print(titles_new)
# df.columns = titles_new
# df = df.drop('配合番号', axis=0)
# print(df)

class Hardness:

    def __init__(self, target):
        self.target = target
        self.exp_name = '硬度_自動集積'
        self.file_path = Service.data_dir(target) + rf'\{self.exp_name}*{target}*.xls*'
        self.file_now = ''

    def FindFile(self):
        print('find files...')
        print(self.file_path)

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)

        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} files(s)')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name}')
            return

    def ReadFile(self):
        print('read file...')
        print(self.file_now)

        df = pd.read_excel(self.file_now, header=7)
        df_target = df.iloc[:,13].to_list()
        print(df_target)

        for i in range(len(df_target)):
            if str(df_target[i]) != 'nan':
                df_target[i] = Service.target_number(i, self.target)
        print(df_target)
        df.iloc[:, 13] = df_target

        df = df.iloc[:,13:16]
        print(df)

        titles = df.columns.to_list()
        print(titles)
        titles = ['配合番号', '０秒','3秒']
        df.columns = titles

        df.dropna(how='all', inplace=True)
        df.dropna(how='all',axis=1,  inplace=True)
        print(df)

        df = df.transpose()
        print(df)
        titles_new = df.loc['配合番号'].to_list()
        print(titles_new)
        df.columns = titles_new
        df = df.drop('配合番号', axis=0)
        print(df)

        print(self.file_now)
        file_name = os.path.splitext(os.path.basename(self.file_now))
        print(file_name)

        print(len(df))

        unit = ['HA'] * len(df)
        type = []
        if len(df) == 1 :
            type = ['０秒']
        elif len(df) == 2:
            type == ['０秒', '3秒']
        condition = [file_name[0]] * len(df)
        method = ['auto tension'] * len(df)

        df.insert(0, 'unit', unit)
        df.insert(0, 'type', type)
        df.insert(0, 'condition', condition)
        df.insert(0, 'method', method)
        df.reset_index(inplace= True, drop= True)

        print(df)

        self.WriteDate(df)


    def WriteDate(self, df_input):
        print('writing data...')

        file_data = Service.data_dir(self.target) + fr'\{self.target} Data.xlsx'
        is_file = os.path.isfile(file_data)


        if is_file:
            pass
        else:
            print('no data file')
            # or you can make a data file
            return

        df_data = pd.read_excel(file_data, index_col=0)
        print(df_data)

        df_merge = pd.concat([df_data, df_input], sort=False)

        df_merge.to_excel(file_data, index=True, header=True)
        print(df_merge)

        print(f'saved data file in {file_data}')


def DoIt(target:str):
    hado = Hardness(target)
    hado.FindFile()

if __name__ == '__main__':
    target = input('target: ')
    DoIt(target)
    