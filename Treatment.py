import pandas as pd
import os
import numpy as np
import Service

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
        print('rounding')
        df = df.round(1)
        print(df)

        print(df[(df['type'] == 'elongation')])
        df[(df['type'] == 'elongation')] = df[(df['type'] == 'elongation')].round(-1)
        print(df)

        ## drop angles of autotension
        # print(df.query('condition in ["Normalアングル", "スチームアングル"] and type in ["25%M", "50%M"]'))
        print(df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'elongation']" ,engine='python'))
        print(df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'elongation']" ,engine='python'))
        index_drop = df.query("condition.str.contains('ｱﾝｸﾞﾙ') and type in ['25%M', '50%M', '100%M', 'elongation']" ,engine='python').index
        df.drop(list(index_drop), inplace=True)

        print(df)



        ## save
        df.to_excel(self.file, index=True, header=True, startcol=0)
        print(f'saved done {self.file}')


def DoIt(target: str):
    toritori = Treatment(target)
    # toritori.ChangeTitles()
    toritori.RoundData()

if __name__ == "__main__":
    print(pd.__version__)
    target = input('target: ')
    DoIt(target)