import os
import glob
import pandas as pd
import Service


class Muuni:

    def __init__(self, target):
        self.exp_name = 'ムーニー_ロータ_自動集積'

        self.target = target
        self.path_xlsx = Service.data_dir(
            target) + rf'\{self.exp_name} {target}*.xlsx'
        self.file_now = ''
        self.index_name = ''

    def FindFile(self):
        print('find files...')
        print(self.path_xlsx)

        file_list = glob.glob(self.path_xlsx)
        file_list = sorted(file_list, key=len)
        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s) ')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name}')
            return

    def ReadFile(self):
        print('read file...', self.file_now)

        df = pd.read_excel(self.file_now, header=None)
        # print(df)
        for i, value in enumerate(df[0]):
            if value == '特性値：':
                num_target = df[0][i-1]
                row_init = i + 4

        print(f'number of target: {num_target}')
        print(f'samples start row: {row_init}')

        df_input = df.loc[[row_init]]
        for i in range(1, num_target):
            df_input = df_input.append(df.loc[[row_init + 2*i]])

        # print('translate row and col')
        df_input = df_input.transpose()
        # print(df_input)

        # print('')
        df_input = df_input.loc[[2, 3, 4]]
        # print(df_input)

        # target numbers
        # print('target numbering')
        # print(len(df_input.columns))
        target_titles = []
        for i in range(len(df_input.columns)):
            target_titles.append(Service.target_number(i, self.target))
        # print(target_titles)
        df_input.columns = target_titles

        # print(df_input)

        # insert unit, method.. else
        file_name = os.path.splitext(os.path.basename(self.file_now))[0]
        print(file_name)
        unit = ['M', 'M', 'min']
        type_list = ['M1', 'Vm', 'T1']
        condition_list = [Service.file_name_without_target_and_expname(
            self.file_now, self.target, self.exp_name)]*len(df_input)
        method_list = [Service.file_name_without_target(
            self.file_now, self.target)]*3

        # print(method_list)
        # print(condition_list)

        df_input.insert(0, 'unit', unit)
        df_input.insert(0, 'type', type_list)
        df_input.insert(0, 'condition', condition_list)
        df_input.insert(0, 'method', method_list)

        print(df_input)

        # return
        self.WriteData(df_input)

    def WriteData(self, df_input):
        print('writing data...')

        file_data = Service.data_dir(
            self.target) + fr'\{self.target} Data.xlsx'

        is_file = os.path.isfile(file_data)

        if is_file:
            pass
        else:
            print('no data file')
            # or you can make a data file
            return

        Service.save_to_data_excel(file_data, df_input, self.exp_name)


def DoIt(target: str):
    muni = Muuni(target)
    try:
        muni.FindFile()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = target = input('target number (ex: ABC001): ')
    DoIt(target)
