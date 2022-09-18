from ast import Del
import glob
import pandas as pd
import numpy as np

class Delta:

    DesktopPath = r'C:\Users\junsa\Desktop'

    def __init__(self, target: str):
        self.target = target
        self.destination = self.DesktopPath + fr'\{target} Data'
        self.filePath = self.destination + fr'\Delta {target}.xlsx'
    
    def ReadFile(self):
        df = pd.read_excel(self.filePath, header=None)
        # df = pd.read_excel(self.filePath, names=['硬度'])
        df = df.fillna(method='ffill')
        # data = df[[1,2,3]]
        data_median = df[[3]].loc[[5, 9, 13, 17, 21]].values
        print(df)

        data_input = ['delta', 'ma','cm']
        print(len(data_median))
        for i in range(len(data_median)):
            print(data_median[i][0])
            data_input.append(data_median[i][0])
        # print(data_median[0][0])
        # print(data_median[1][0])
        # print(data_input)
        self.InputData(data_input)
    
    def InputData(self, data_input):
        print('inputing data....')
        print(data_input)

        excelPath = self.destination + fr'\{self.target} Data.xlsx'
        df_copy = pd.read_excel(excelPath, index_col=0)
        df_copy.reset_index(inplace= True, drop= True)
        df_copy.loc[len(df_copy)] =data_input
        print(df_copy)

        df_copy.to_excel(excelPath, index=True, header=True, startcol=1)
        


        

def Del(target: str, DesktopPath: str):
    file_list = glob.glob(DesktopPath + rf'\{target} Data\Delta {target}*.xlsx')
    if len(file_list) > 0 :
        print()
        print()
        print('Delta file exist')
        print(file_list)
        delta = Delta(target)
        delta.ReadFile()
    else:
        print('no Delta file')


if __name__ == 'Delta.py':
    Delta()




