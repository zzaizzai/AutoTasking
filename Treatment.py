from turtle import title
import pandas as pd
import os


class Treatment:

    def __init__(self, target):
        self.target = target
        self.desktop = os.path.expanduser('~/Desktop')
        self.file = self.desktop + \
            rf'\{self.target} Data\{self.target} Data.xlsx'


    def ChangeTitles(self):
        
        is_file = os.path.isfile(self.file)
        if is_file:
            pass
        else:
            print('no data file')
            return

        def target_number(number: int):
            target = self.target
            alphabet = target[0:3]
            num = int(target[3:])
            alphabet_num = alphabet + str('%03d' % (num + number))
            return alphabet_num

        df = pd.read_excel(self.file, index_col=0)
        print(df)

        titles = ['method', 'condition', 'type', 'unit']

        for i in range(len(df.columns[4:])):
            titles.append(target_number(i))

        df.columns = titles

        print(df)
        # df.to_excel(self.file, index=True, header=True, startcol=0)
        # print(f'saved done {self.file}')
        


def toriri(target: str):
    toritori = Treatment(target)
    toritori.ChangeTitles()


if __name__ == "__main__":
    target = input('target: ')
    toriri(target)
