import matplotlib.pyplot as plt
import pandas as pd
import os


class Gurapu:
    def __init__(self, target):
        self.desktop = os.path.expanduser('~/Desktop')
        self.target = target
        self.file = self.desktop + rf'\{self.target} Data\レオメータ CBA001-007 .xlsx'
        
    def MakeGraph(self):
        df = pd.read_excel(self.file, header=None)
        print(df)

        row_start = 0
        for i, value in enumerate(df[0]):
            print(value)
            if value == 'Time(NO.1)':
                row_start = i
                break
        print(row_start
        )
        df = df[row_start:]
        df = df.dropna(axis=1)
        print(df)
        df.columns = df.iloc[0]
        df = df[1:]
        df.reset_index(inplace=True, drop=True)
        df = df.T.reset_index(drop=True).T
        print(df)

        df.plot(0)
        plt.xlim(0,)
        plt.ylim(0,)
        plt.xlabel('time (min)')
        plt.ylabel('sss')
        plt.grid()
        plt.show()
        # df = df[21:]
        # df = df[[0,1,5,9,13,17]]
        # print(df)

        # df.reset_index(inplace= True, drop= True)
        # df = df.T.reset_index(drop=True).T

        # print(df)


        # plt.rcParams['font.family'] = 'MS Gothic'
        # plt.rcParams['font.size'] = 14

        # time = df[0].values
        # aa = df[1].values
        # bb = df[2].values

        # plt.plot(time, aa)
        # plt.plot(time, bb)
        # plt.legend(['aa', 'cc'])
        # plt.grid()
        # plt.show()

def guragura(target: str):
    gura = Gurapu(target)
    gura.MakeGraph()

if __name__ == '__main__':
    print('make a graph')

    target = 'CBA001'
    guragura(target)
    