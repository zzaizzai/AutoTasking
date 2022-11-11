import os
import glob
import pandas as pd
import Service


class Muuni:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target):
        self.exp_name = 'ムーニー_ロータ_自動集積'

        self.target = target
        self.path_xlsx = Service.data_dir(
            target) + rf'\{self.exp_name} {target}*.xlsx'
        self.file_now = ''
        self.index_name = ''

    def FindFile(self):
        print('find files as ', os.path.basename(self.path_xlsx))
        # print(self.path_xlsx)

        file_list = glob.glob(self.path_xlsx)
        file_list = sorted(file_list, key=len)
        # print(file_list)

        df_all = pd.DataFrame()
        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s) ')

            for file in file_list:
                self.file_now = file
                df_all = pd.concat([df_all, self.ReadFile()])
        else:
            print(f'No {self.exp_name}')
            return

        df_all = self.RemoveOtherInfo(df_all)
        
        self.WriteData(df_all)

    def RemoveOtherInfo(self, df_all):
        print(df_all)

        df_all = df_all.reset_index(drop=True)
        df_all_temp_condition = df_all["condition"]

        index_remove = []
        for i, value in enumerate(df_all_temp_condition):
            if ("121℃" not in str(value)) and ("Vm" in df_all["type"][i] or "5p" in df_all["type"][i]):
                print(value)
                index_remove.append(i)
        df_all = df_all.drop(index=index_remove )
        return df_all

    def ReadFile(self):
        print('read file...', os.path.basename(self.file_now))

        df = pd.read_excel(self.file_now, header=None)

        for i, value in enumerate(df[0]):
            if value == '特性値：':
                num_target = df[0][i-1]
                row_init = i + 4

        if self.test_mode:
            print(f'number of target: {num_target}')
            print(f'samples start row: {row_init}')

        df_input = df.loc[[row_init]]
        for i in range(1, num_target):
            df_input = df_input.append(df.loc[[row_init + 2*i]])


        df_input = df_input.transpose()

        df_input = df_input.loc[[2, 3, 4]]


        target_titles = []
        for i in range(len(df_input.columns)):
            target_titles.append(Service.target_number(i, self.target))
        # print(target_titles)
        df_input.columns = target_titles


        file_name = os.path.splitext(os.path.basename(self.file_now))[0]

        unit = ['kgf・cm', 'kgf・cm', 'min']
        type_list = ['MV', 'Vm', 'ST 5p']

        condition_teperature = df.iat[1,5]
        condition_teperature = str(int(float(condition_teperature.split("℃")[0]))) + '℃'

        condition_list = [condition_teperature]*len(df_input)

        method_list = [Service.file_name_without_target(
            self.file_now, self.target)]*3



        df_input.insert(0, 'unit', unit)
        df_input.insert(0, 'type', type_list)
        df_input.insert(0, 'condition', condition_list)
        df_input.insert(0, 'method', method_list)

        if self.test_mode:
           print(df_input)

        # return
        return df_input
        # self.WriteData(df_input)

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


def DoIt(target: str, test_mode = False):
    muni = Muuni(target)
    muni.TestMode(test_mode)

    try:
        muni.FindFile()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = target = input('target number (ex: ABC001): ')
    DoIt(target, test_mode = True)
