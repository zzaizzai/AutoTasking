import os
import glob
import pandas as pd
import Service


class Muuni:

    def __init__(self, target):
        self.test_mode = False
        
        self.exp_name = 'ムーニー_ロータ_自動集積'
    
        self.target = target
        self.path_xlsx = Service.data_dir(
            target) + rf'\{self.exp_name} {target}*.xlsx'
        self.file_now = ''
        self.index_name = ''

    def set_testmode(self, mode: bool):
        self.test_mode = mode

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

        df_all = self.Change5pValue(df_all)

        self.WriteData(df_all)

    def Change5pValue(self, df_all: pd.DataFrame) -> pd.DataFrame:

        df_temp = df_all

        if self.test_mode:
            print(df_temp)
        for _, col_title in enumerate(df_temp.columns.values[4:]):
            for row_index in df_temp[col_title].index.to_list():
                if df_temp.at[row_index, col_title] == '******':
                    df_temp.at[row_index, col_title] = "30 >"

        # print(df_temp)
        return df_temp
        

    def RemoveOtherInfo(self, df_all: pd.DataFrame):
        if self.test_mode == True:
            print(df_all)

        df_all = df_all.reset_index(drop=True)
        df_all_temp_condition = df_all["condition"]

        index_remove = []
        for i, value in enumerate(df_all_temp_condition):
            if ("121℃" not in str(value)) and ("Vm" in df_all["type"][i] or "5p" in df_all["type"][i]):
                print(value)
                index_remove.append(i)
        df_all = df_all.drop(index=index_remove)
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
        df_input.columns = target_titles


        unit_list = ['kgf・cm', 'kgf・cm', 'min']
        type_list = ['MV', 'Vm', 'ST 5p']

        condition_name = df.iat[1, 5]
        condition_name = str(
            int(float(condition_name.split("℃")[0]))) + '℃'
        condition_from_file_name = Service.file_name_without_target_and_expname_distin_underbar(self.file_now, self.target, self.exp_name)
        print('file name:',condition_from_file_name)

        if condition_from_file_name != "none" and len(condition_from_file_name) > 5:
            condition_name = condition_name +  " " + condition_from_file_name

        condition_list = [condition_name]*len(df_input)
        method_list = [Service.file_name_without_target_distin_underbar(self.file_now, self.target)]*3

        df_input = Service.create_method_condition_type_unit(df_input, method_list, condition_list, type_list, unit_list)

        if self.test_mode:
            print(df_input)

        return df_input

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


def DoIt(target: str, test_mode=False):
    muni = Muuni(target)
    muni.set_testmode(test_mode)

    try:
        muni.FindFile()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = target = input('target number (ex: ABC001): ')
    DoIt(target, test_mode=True)
