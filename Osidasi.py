from operator import index
from matplotlib.pyplot import title
import pandas as pd
import Service
import glob


class Osidasi:

    def __init__(self, target: str):
        self.exp_name = '押出し '
        self.target = target
        self.file_path = Service.data_dir(target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_data = Service.data_dir(target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''
    
    def StartProcess(self):
        print('find file')
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
        print('Read file')


        
        df = pd.read_excel(self.file_now, header=16, index_col=1)
        print(df)
        print(df.index.to_list())

        target_list = df.index.to_list()
        print(len(target_list))
        for i, value in enumerate(target_list):
            print(i , value)
            if str(value) == 'nan' and  str(target_list[i-1]) != 'nan' and str(target_list[i+1]) == 'nan' and i > 0  and i < len(target_list) -3:
                print('it is next position of target')
                target_list[i] = target_list[i-1]
        print(target_list)
        df.index = target_list


        target_list_set = set(target_list)
        target_list_numbers = list(target_list_set)

        # without nan
        number_of_target:int = len(target_list_numbers) - 1
        print('number_of_target',number_of_target)

        df = df.iloc[:number_of_target*4+1,:]
        print(df)



        print(df.columns.to_list())
        titles = df.columns.to_list()
        titles[0] = 'mean'
        df.columns = titles
        print(len(df))
        
        # mean data 
        # print(target_list)
        count_for_mean = 0
        index_mean = []
        index_eval = []
        for i in range(len(target_list)):
            print(target_list[i])
            if str(target_list[i]) != 'nan':
                count_for_mean += 1
                if count_for_mean % 3 == 0:
                    print('mean')
                    index_mean.append(i)
                    index_eval.append(i-2)
                if count_for_mean > number_of_target*3:
                    break
        print('index_mean',index_mean)
        print('index_eval',index_eval)

        # get only mean data
        df_mean = df.iloc[index_mean,:]
        print(df_mean.loc[:,['L','W','Swell','Swell.1']])


        # get evaluations data
        df_eval = df.iloc[index_eval,:]
        print(df_eval.loc[:,['H.1','Sc','R.F']])
def DoIt(target:str):
    osiosi = Osidasi(target)
    try:
        osiosi.StartProcess()
    except Exception as e:
        print(e)


if __name__ == "__main__":
    target = input('target: ')
    DoIt(target)