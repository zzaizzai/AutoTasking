import pandas as pd
import os
import numpy as np
import Service
import openpyxl

class Treatment:

    def __init__(self, target):
        self.target = target
        self.file = Service.data_dir(target) +  rf'\{self.target} Data.xlsx'

    def ChangeTitles(self):
        
        is_file = os.path.isfile(self.file)
        if is_file:
            pass
        else:
            print('no data file')
            return

        df = pd.read_excel(self.file, index_col=0)
        # print(df)

        titles = ['method', 'condition', 'type', 'unit']

        for i in range(len(df.columns[4:])):
            titles.append(Service.target_number(i))

        df.columns = titles

        # print(df)
        df.to_excel(self.file, index=True, header=True, startcol=0)
        print(f'saved done {self.file}')

    def RoundData(self):

        is_file = os.path.isfile(self.file)
        
        if is_file:
            pass
        else:
            print('no data file')
            return

        print('round data')
        df = pd.read_excel(self.file, index_col=0)
        

        print(df)
        print('before rounding')

        print('df lengh: ', len(df))

        if len(df) < 1:
            print('no df')
            return
        else:
            pass

        ## round
        df = df.replace('******', 0)
        print('replaced ****** to 0')

        df = df.astype('float', errors='ignore')
        
        df = Service.normal_round(df, 1)

        df[(df['type'] == '破断伸び％')] = Service.normal_round(df[(df['type'] == '破断伸び％')], -1)
        df[(df['type'] == 'EB')] = Service.normal_round(df[(df['type'] == 'EB')], -1)
        df[(df['type'] == 'elongation')] = Service.normal_round(df[(df['type'] == 'elongation')], -1)


        df[(df['type'] == '０秒')] = Service.normal_round(df[(df['type'] == '０秒')],0)
        df[(df['type'] == '3秒')] = Service.normal_round(df[(df['type'] == '3秒')],0)
        df[(df['type'] == 'HA(0s)')]= Service.normal_round(df[(df['type'] == 'HA(0s)')],0)
        df[(df['type'] == '⊿V')]= Service.normal_round(df[(df['type'] == '⊿V')],0)
        # print(df)

        print('rounding done')
        print(df)

        ## drop angles of autotension
        # print(df.query('condition in ["Normalアングル", "スチームアングル"] and type in ["25%M", "50%M"]'))
        # print(df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'elongation']" ,engine='python'))
        # print(df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'EB']" ,engine='python'))
        index_drop = df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'EB']" ,engine='python').index
        df.drop(list(index_drop), inplace=True)

        # print(df)



        ## save
        df.to_excel(self.file, index=True, header=True, startcol=0)
        print(f'saved done {self.file}')


    def CellWidth(self):
        print('fixing cell width')
        
        wb = openpyxl.load_workbook(self.file)
        ws = wb.worksheets[0]

        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 20

        wb.save(self.file)


    def Sorting(self):
        print('sorting')
        
        is_file = os.path.isfile(self.file)
        
        if is_file:
            pass
        else:
            print('no data file')
            return
        df = pd.read_excel(self.file, index_col=0)
        
def DoIt(target: str):
    toritori = Treatment(target)
    try:
        # toritori.ChangeTitles()
        toritori.RoundData()
        # toritori.Sorting()
        toritori.CellWidth()
    except Exception as e:
        print(e)

if __name__ == "__main__":
    print(pd.__version__)
    target = input('target: ')
    DoIt(target)