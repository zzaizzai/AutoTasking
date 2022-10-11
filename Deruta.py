import pandas as pd
import Service
import glob


class Deruta:

    def __init__(self, target):
        self.exp_name = '⊿Ｖ'

        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_now = ''

    def FindFile(self):
        print('find files....')

        print(self.file_path)

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)
        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list) } {self.exp_name} file(s)')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name}')
            return

    def ReadFile(self):
        print('read file...')

        print(self.file_now)

        df = pd.read_excel(self.file_now, sheet_name='1')

        new_col = df.columns.to_list()
        print(df.columns.to_list())
        new_col[2] = '配合番号'
        new_col[3] = 'liquid_index'
        new_col[4] = 'liquid'
        new_col[7] = 'condition_index'
        new_col[8] = 'condition'
        df.columns = new_col
        print(df)

        df = df.iloc[:, 2:]
        print(df)

        # count number of target
        target_list = df['配合番号'].values.tolist()
        print(target_list)

        # get targets
        target_list_temp = []
        for i in range(len(target_list)):
            # print(target_list[i])
            if str(target_list[i]) != 'nan':
                target_list_temp.append(target_list[i])

        target_list = target_list_temp
        target_list_set = set(target_list)
        target_list = list(target_list_set)
        print(target_list)

        conditions_list_index = df.query(
            "liquid_index in ['試験液']").index.to_list()
        print(conditions_list_index)

        condition_list = []
        for i in conditions_list_index:
            condition_list.append(
                str(df.loc[:, 'liquid'][i]) + " " + str(df.loc[:, 'condition'][i]))
        print(condition_list)

        df_all = pd.DataFrame()
        for i in range(len(condition_list)):
            df_all = pd.concat([df_all, self.ReadDataBlock(
                condition_list[i], conditions_list_index[i], len(target_list))])

        print(df_all)

    def ReadDataBlock(self, condition_of_exp, conditions_list_index: int, numbers_target: int):

        print(
            f'read data block with {condition_of_exp}, index: {conditions_list_index}, targets: {numbers_target}')

        df = pd.read_excel(self.file_now, sheet_name='1',
                           header=conditions_list_index+4, index_col=2)
        df = df.iloc[:3*numbers_target, :]
        print(df)
        index_temp = df.index.to_list()
        for i in range(1, len(index_temp) - 1):
            if str(index_temp[i]) != 'nan' and str(index_temp[i-1]) == 'nan' and str(index_temp[i+1]) == 'nan':
                index_temp[i-1] = index_temp[i]
                index_temp[i+1] = index_temp[i]
        print(index_temp)

        index_target_temp = []
        for value in index_temp:
            index_target_temp.append(
                Service.target_number(value-1, self.target))

        print(index_target_temp)

        df.index = index_target_temp
        print(df)

        # change col titles
        print('change col titles')
        titles_temp = df.columns.values.tolist()
        titles_temp[2] = 'mean'
        df.columns = titles_temp
        print(df)

        df = df.query("mean in ['平均値']")
        print(df)

        df = df.loc[:, ['⊿V']]
        print(df)

        df = df.transpose()

        condition = [condition_of_exp]
        method = ['oil']
        unit = ['??']
        type = ['dertua']

        df.insert(0, 'unit', unit)
        df.insert(0, 'type', type)
        df.insert(0, 'condition', condition)
        df.insert(0, 'method', method)

        df.reset_index(inplace=True, drop=True)
        print(df)

        return df

def DoIt(target: str):
    ruta = Deruta(target)
    ruta.FindFile()


if __name__ == '__main__':
    target = input('target: ')

    DoIt(target)