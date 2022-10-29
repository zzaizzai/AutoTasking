import pandas as pd
import Service
import glob
import os


class Deruta:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target):
        self.exp_name = '⊿Ｖ'

        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'
        self.file_data = Service.data_dir(
            target) + fr'\{self.target} Data.xlsx'
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

        df = pd.read_excel(self.file_now, sheet_name='1',
                           index_col=0, header=0)
        # print(df)
        new_col = df.columns.to_list()
        # print(df.columns.to_list())
        new_col[1] = '配合番号'
        new_col[2] = 'liquid_index'
        new_col[3] = 'liquid'
        new_col[7] = 'condition_index'
        new_col[7] = 'condition'
        new_col[8] = 'condition_time'
        df.columns = new_col

        print('after rename of title')
        # print(df)

        # count number of target
        target_list = df['配合番号'].values.tolist()
        print('target list', target_list)

        # get targets
        target_list_temp = []
        for i in range(len(target_list)):
            # print(target_list[i])
            if not str(target_list[i]).isalpha() and str(target_list[i]) != 'nan':
                # print(target_list[i])
                target_list_temp.append(int(target_list[i]))
        print('target temp', target_list_temp)

        target_list = target_list_temp
        target_list_set = set(target_list)
        target_list = list(target_list_set)
        print('target list', target_list)

        # generate condition list

        # return
        conditions_list_index = []
        print(df['liquid_index'])
        for i, value in enumerate(df['liquid_index']):
            if value == "試験液":
                conditions_list_index.append(int(i))
        print('condition list index', conditions_list_index)

        print('??')
        condition_list = []
        for index_liquid in conditions_list_index:
            # print('index of liquid',index_liquid)
            condition_name = str(df.iat[index_liquid, 3]) + ' ' + str(
                df.iat[index_liquid, 7]) + '℃×' + str(df.iat[index_liquid, 8])
            condition_list.append(condition_name)
        print('condition list', condition_list)

        df_all = pd.DataFrame()
        for i in range(len(condition_list)):
            df_all = pd.concat([df_all, self.ReadDataBlock(
                condition_list[i], conditions_list_index[i], len(target_list))])
        # df_all.iloc[:,4:] = df_all.iloc[:,4:].round(0)
        # print(df_all.iloc[:,4:].round(1))
        print('???')
        print(df_all)

        self.writedata(df_all)

    def ReadDataBlock(self, condition_of_exp, conditions_list_index: int, numbers_target: int):

        print(
            f'read data block with {condition_of_exp}, index: {conditions_list_index}, targets: {numbers_target}')

        df = pd.read_excel(self.file_now, sheet_name='1',
                           header=conditions_list_index+4, index_col=2)
        df = df.iloc[:3*numbers_target, :]
        # print(df)
        index_temp = df.index.to_list()
        for i in range(1, len(index_temp) - 1):
            if str(index_temp[i]) != 'nan' and str(index_temp[i-1]) == 'nan' and str(index_temp[i+1]) == 'nan':
                index_temp[i-1] = index_temp[i]
                index_temp[i+1] = index_temp[i]
        # print(index_temp)

        df.index = index_temp

        # remove nan index
        df = df.query("index == index")
        index_temp = df.index.to_list()
        # print(df)

        index_target_temp = []
        for value in index_temp:
            target_inddex = int(value) - int(self.target[3:])
            index_target_temp.append(
                Service.target_number(target_inddex, self.target))

        # print(index_target_temp)

        df.index = index_target_temp
        # print(df)

        # change col titles
        print('change col titles')
        titles_temp = df.columns.values.tolist()
        titles_temp[2] = 'mean'
        df.columns = titles_temp
        # print(df)

        df = df.query("mean in ['平均値']")
        # print(df)

        df = df.loc[:, ['△Ｖ']]
        # print(df)

        df = df.transpose()

        condition = [condition_of_exp]
        # method = ['oil']
        method = [Service.file_name_without_target(self.file_now, self.target)]
        unit = ['%']
        type_list = ['⊿V']

        df.insert(0, 'unit', unit)
        df.insert(0, 'type', type_list)
        df.insert(0, 'condition', condition)
        df.insert(0, 'method', method)

        df.reset_index(inplace=True, drop=True)

        # print(df.iloc[:,4:].round(1))
        if self.test_mode:
            print(df)

        return df

    def writedata(self, df_input):
        print('writing data')

        is_file = os.path.isfile(self.file_data)

        if is_file:
            pass
        else:
            print('no data file')
            return

        Service.save_to_data_excel(self.file_data, df_input, self.exp_name)


def DoIt(target: str, test_mode=False):
    ruta = Deruta(target)
    ruta.TestMode(mode=test_mode)

    try:
        ruta.FindFile()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = input('target: ')

    DoIt(target, test_mode=True)
