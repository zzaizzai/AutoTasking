import pandas as pd
import os
import numpy as np
import Service
import openpyxl


class Treatment:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target):
        self.target = target
        self.file = Service.data_dir(target) + rf'\{self.target} Data.xlsx'

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

        print('rounding data')
        df = pd.read_excel(self.file, index_col=0)

        if self.test_mode:
            print('before rounding')
            print(df)

        # print('df lengh: ', len(df))

        if len(df) < 1:
            print('no df')
            return
        else:
            pass

        # type 0.1
        df = Service.round_by_check_eachone(df, 1, [
            'ST 5p', 'ML', 'MV', 'T1', 'MH', 'ts1', 't10', 't50', 't90', 'CR',
            'M25', 'M50', 'M100', 'TS', 'Tr-B'
        ])
        # type 0
        df = Service.round_by_check_eachone(
            df, 0, ['０秒', '3秒', 'HA(0s)', 'HA(3s)', '⊿V'])

        # type 10
        df = Service.round_by_check_eachone(df, -1,
                                            ['elongation', 'EB', '破断伸び％'])

        if self.test_mode:
            print('rounding done')
            print(df)

        index_drop = df.query(
            "condition.str.contains('ｱﾝｸﾞﾙ') and type in ['M25', 'M50', 'M100', 'EB']",
            engine='python').index
        df.drop(list(index_drop), inplace=True)

        # print(df)

        # save
        df.to_excel(self.file, index=True, header=True, startcol=0)
        print('data rounding done')
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


def DoIt(target: str, test_mode=False):
    toritori = Treatment(target)
    toritori.TestMode(test_mode)
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
