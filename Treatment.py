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
        print(df)

        titles = ['method', 'condition', 'type', 'unit']

        for i in range(len(df.columns[4:])):
            titles.append(Service.target_number(i))

        df.columns = titles

        print(df)
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

        print('df lengh: ', len(df))

        if len(df) < 1:
            print('no df')
            return
        else:
            pass

        ## round\
        df = df.replace('******', 0)
        # print('rounding')

        df = df.astype('float', errors='ignore')
        
        df = df.round(1)
        print(df)

        print(df[(df['type'] == 'elongation')])
        df[(df['type'] == 'elongation')] = df[(df['type'] == 'elongation')].round(-1)
        df[(df['type'] == 'EB')] = df[(df['type'] == 'EB')].round(-1)
        df[(df['type'] == '破断伸び％')] = df[(df['type'] == '破断伸び％')].round(-1)
        df[(df['type'] == '０秒')] = df[(df['type'] == '０秒')].round(0)
        df[(df['type'] == '3秒')] = df[(df['type'] == '3秒')].round(0)
        df[(df['type'] == 'HA(0s)')] = df[(df['type'] == 'HA(0s)')].round(0)
        df[(df['type'] == '⊿V')] = df[(df['type'] == '⊿V')].round(0)
        print(df)

        ## drop angles of autotension
        # print(df.query('condition in ["Normalアングル", "スチームアングル"] and type in ["25%M", "50%M"]'))
        # print(df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'elongation']" ,engine='python'))
        print(df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'EB']" ,engine='python'))
        index_drop = df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'EB']" ,engine='python').index
        df.drop(list(index_drop), inplace=True)

        print(df)



        ## save
        df.to_excel(self.file, index=True, header=True, startcol=0)
        print(f'saved done {self.file}')


    def CellWidth(self):
        print('cell width...')
        
        wb = openpyxl.load_workbook(self.file)
        ws = wb.worksheets[0]

        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 15

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