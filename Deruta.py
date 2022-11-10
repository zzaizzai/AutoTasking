import pandas as pd
import Service
import glob
import os
import numpy


class Deruta:

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target):
        self.test_mode = False

        self.exp_name = 'ΔV'

        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'
        self.file_data = Service.data_dir(
            target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''

    def FindFile(self):
        print(f'find files of {self.exp_name}')

        if self.test_mode:
            print(self.file_path)

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)

        if self.test_mode:
            print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list) } {self.exp_name} file(s)')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name}')
            return

    def FindIndexOfTitle(self, df) -> int:
        # print(df.columns.to_list().index('配合ＮＯ'))
        # print(df.columns.to_list())
        target_col_index = None
        for index_title, title in enumerate(df.columns.to_list()):
            for index, cell in enumerate(df[title]):
                # print(index, cell)
                if '試験液' in str(cell):
                    # print(title, index)
                    target_col_index = index_title
                    break
        return target_col_index - 1

    def ReadFile(self):
        print('read file ', os.path.basename(self.file_now))

        df = pd.read_excel(self.file_now, sheet_name='1',
                           index_col=None, header=None)

        target_col_index = self.FindIndexOfTitle(df)
        print('target col : ', target_col_index)

        new_col = df.columns.to_list()
        # print(df.columns.to_list())
        # print(df)

        liquid_col_index = target_col_index + 1
        liquid_col_kind = liquid_col_index + 1
        condition_index = target_col_index + 6
        condition_kind = condition_index + 1
        condition_time = condition_index + 2

        new_col[target_col_index] = '配合番号'
        new_col[liquid_col_index] = 'liquid_index'
        new_col[liquid_col_kind] = 'liquid'

        new_col[condition_index] = 'condition_index'
        new_col[condition_kind] = 'condition'
        new_col[condition_time] = 'condition_time'
        df.columns = new_col

        print(df)

        # print('after rename of title')
        # print(df)

        # count number of target
        target_list = df['配合番号'].values.tolist()

        # print(target_list)

        for index, value in enumerate(target_list):
            # print(value)
            if "プレス1次" in str(value) or "スチーム1次" in str(value):
                target_list[index] = 'nan'
        # print('target list', target_list)

        # # get targets
        target_list_temp = [x for x in target_list if str(
            x) != 'nan' and str(x).replace(".", "", 1).isdigit()]
        # print(target_list_temp)
        target_list_index = [i for i, x in enumerate(target_list) if str(
            x) != 'nan' and str(x).replace(".", "", 1).isdigit()]
        target_list = list(set(target_list_temp))
        # if self.test_mode:
        print('showginf target list')
        print(target_list)
        print(target_list_index)

        # print(target_list)
        if self.test_mode:
            print('target list', target_list)

        conditions_list_index = [int(i) for i, value in enumerate(
            df['liquid_index']) if value == "試験液"]
        # print(df['liquid_index'])
        print('condition list index', conditions_list_index)

        def get_condition_list(row_index_condition):
            print(f'get condition lsit row of {row_index_condition}')
            condition_name = str(df.iat[row_index_condition, liquid_col_kind]) + ' ' + str(
                df.iat[row_index_condition, condition_index]) + '℃×'
            if str(str(df.iat[row_index_condition, condition_index + 1])) != 'nan':
                condition_name = condition_name + \
                    str(df.iat[row_index_condition, condition_index + 1])

            if str(df.iat[row_index_condition + 3, condition_index + 1]) != "nan":
                condition_name = condition_name + \
                    str(df.iat[row_index_condition + 1, condition_index + 3])
            print(condition_name)
            return condition_name

        # condition_list = []
        # for row_index_liquid in conditions_list_index:
        #     # print('index of liquid',index_liquid)
        #     condition_name = str(df.iat[row_index_liquid, liquid_col_kind]) + ' ' + str(
        #         df.iat[row_index_liquid, condition_index]) + '℃×' + str(df.iat[row_index_liquid, condition_index + 1])
        #     condition_list.append(condition_name)
        # print('condition list', condition_list)

        df_all = pd.DataFrame()
        condition_block_index = 0

        for i in range(0, len(target_list_index), len(target_list)):
            print('all condition index', conditions_list_index)
            condition_candidate = target_list_index[i: i +
                                                    len(target_list)][0] - 5
            print(f'condition_candidate {condition_candidate}')
            if condition_candidate in conditions_list_index:
                print(f'condition index {condition_candidate} is in list ')
                condition_block_index = condition_candidate
            else:
                print(
                    f'might be it is same to the before block {condition_candidate} altinatly use index: {condition_block_index}')
            condition_block_value = get_condition_list(condition_block_index)
            df_block_for_merge = self.ReadBlock_2(
                df, target_list, target_list_index[i: i+len(target_list)], condition_block_value)
            df_all = pd.concat([df_all, df_block_for_merge], sort=False)
        # self.ReadBlock_2(df, target_list, target_col_index)
        # for i in range(len(condition_list)):
        #     df_all = pd.concat([df_all, self.ReadDataBlock(
        #         condition_list[i], conditions_list_index[i], len(target_list))])

        print(df_all)
        self.writedata(df_all)

        # self.writedata(df_all)
    def ReadBlock_2(self, df, target_list: list, target_list_index: list, condition: str):
        print('targets', target_list, 'target index', target_list_index)
        # print(df)
        df_block = df.iloc[target_list_index[0]-2:target_list_index[-1]+2, :]

        df_block.reset_index(inplace=True, drop=True)
        titles_new = df_block.iloc[0, :].to_list()
        df_block = df_block.drop(index=[0])
        df_block.reset_index(inplace=True, drop=True)

        for i, title in enumerate(titles_new):
            if title == "空中重量":
                titles_new[i - 2] = "配合番号"
                titles_new[i - 1] = "順番"
                break

        df_block.columns = titles_new

        df_block_temp = df_block
        # print(df_block_targets_temp)
        targets_index = [i for i, x in enumerate(
            df_block_temp["配合番号"]) if str(x) != 'nan']
        # print(targets_index)
        for index in targets_index:
            # print(df_block_targets_temp[index])
            # new_value = df_block_targets_temp[index].copy()
            df_block_temp.loc[index-1,
                              '配合番号'] = df_block_temp.loc[index, '配合番号']
            df_block_temp.loc[index+1,
                              '配合番号'] = df_block_temp.loc[index, '配合番号']
            # df_block_targets_temp[index +1] = df_block_targets_temp[index]
        # print(df_block_temp)
        df_block["配合番号"] = df_block_temp["配合番号"]

        # print(df_block.loc[df_block["順番"] == "平均値",["配合番号","⊿V"]])

        df_block_input = df_block.loc[df_block["順番"] == "平均値", ["配合番号", "⊿V"]]
        df_block_input = df_block_input.transpose()

        # print(df_block_input)
        df_block_input.reset_index(inplace=True, drop=True)
        titles_new = [Service.target_number(i, self.target) for i, _ in enumerate(
            df_block_input.iloc[0, :].to_list())]
        df_block_input.columns = titles_new
        df_block_input = df_block_input.drop(index=[0])
        df_block_input.reset_index(inplace=True, drop=True)

        df_block_input.insert(0, 'unit', '%')
        df_block_input.insert(0, 'type', "⊿V")
        df_block_input.insert(0, 'condition', condition)
        df_block_input.insert(0, 'method', Service.file_name_without_target(self.file_now, self.target))

        # print(df_block_input)

        return df_block_input

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
