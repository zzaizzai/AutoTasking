from cgi import test
import pandas as pd
import os
import glob
import Service


class Tension:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target: str):
        self.DesktopPath = os.path.expanduser('~/Desktop')

        self.data_dir = Service.data_dir(target)

        self.auto_file_dir = self.data_dir + r'\auto_tension'
        self.target = target
        self.file_data = self.data_dir + fr'\{self.target} Data.xlsx'
        self.len_df_col = 0

    def StartProcess(self):

        is_file = os.path.isfile(self.file_data)

        if is_file:
            pass
        else:
            print('no data file')
            return

        auto_file_path = self.auto_file_dir + r'\*.xlsm'
        auto_file_list = glob.glob(auto_file_path)

        if len(auto_file_list) == 0:
            print('no auto tesnion data file')
            return

        df_all = pd.DataFrame()

        for auto_file in auto_file_list:
            print(os.path.basename(auto_file))
            df_all = pd.concat([df_all, self.GetData(auto_file)])

        print('showing merge df with sorted')
        # sort witt tile
        df_all = df_all.sort_values(by=[2, 1])

        df_all.reset_index(inplace=True, drop=True)
        df_all = df_all.T.reset_index(drop=True).T

        # remove another targets with checking target number and data file lengh
        df_len = pd.read_excel(self.file_data, index_col=0)
        # without condition, method, type, unit
        self.len_df_col = len(df_len.columns) - 4
        print(f'len of col with only data : {self.len_df_col}')

        if self.len_df_col < 1:
            print('you should do another method first')
            return

        remove_row_list = []

        # get range like df
        if self.test_mode:
            print('all df')
            print(df_all)

        for i, value in enumerate(df_all[0]):
            if int(value[3:]) < int(self.target[3:]):
                if self.test_mode:
                    print(int(value[3:]))
                    print('below range ')
                remove_row_list.append(i)
            if int(value[3:]) >= self.len_df_col + int(self.target[3:]):
                if self.test_mode:
                    print(int(value[3:]))
                    print('over range')
                remove_row_list.append(i)
        df_all = df_all.drop(remove_row_list)

        df_all.reset_index(inplace=True, drop=True)
        df_all = df_all.T.reset_index(drop=True).T

        target_condition_list = []
        for value in df_all[1]:
            target_condition_list.append(value)

        target_condition_set = set(target_condition_list)
        target_condition_list = list(target_condition_set)

        print(target_condition_list)

        target_condition_list = sorted(target_condition_list,
                                       key=len,
                                       reverse=True)

        df_merge = pd.DataFrame()
        for condition in target_condition_list:
            df_part = pd.DataFrame()

            for i, value in enumerate(df_all[1]):
                if value == condition:
                    df_part = df_part.append(df_all.loc[[i]])

            df_part = df_part.transpose()
            df_part = df_part.T.reset_index(drop=True).T
            df_part.reset_index(inplace=True, drop=True)
            target_titles = df_part.iloc[[0]].values.tolist()[0]

            unit = ['MPa', 'MPa', 'MPa', 'MPa', '%']
            method = ['M25', 'M50', 'M100', 'TS', 'EB']

            condition_name = df_part[0][1]
            condition = [condition_name] * 5

            name = ['初期物性'] * 5
            df_part = df_part.loc[[2, 3, 4, 5, 6]]

            df_part.columns = target_titles

            df_part.insert(loc=0, column='unit', value=unit)
            df_part.insert(loc=0, column='type', value=method)
            df_part.insert(loc=0, column='condition', value=condition)
            df_part.insert(loc=0, column='method', value=name)

            df_part.reset_index(inplace=True, drop=True)

            df_part = df_part.loc[:, ~df_part.columns.duplicated(keep="last")]

            df_part = self.ChangeBrTension(df_part)
            df_merge = pd.concat([df_merge, df_part])

        # df_merge = self.drop_anguru_without_ts(df_merge)

        self.WriteData(df_merge)
        # print('merge done')

    # def drop_anguru_without_ts(self, df:pd.DataFrame) -> pd.DataFrame:

    #     df.reset_index(inplace=True, drop=True)
    #     index_drop = df.query(
    #         "condition.str.contains('ｱﾝｸﾞﾙ') and type in ['M25', 'M50', 'M100', 'EB']", engine='python').index
    #     df.drop(list(index_drop), inplace=True)
    #     index_press = df.query("condition.str.contains('PressJIS')").index
    #     print(index_press)
    #     print(df)

    #     df_ordering_unit = pd.DataFrame({'unit_order': ['1', 2,3,4,5,6], 'unit':['M25', 'M50','M100', 'TS', 'EB' ,'Tr-B']})
    #     print(pd.merge(df, df_ordering_unit, on="unit" ,how='left'))
    #     print(df_ordering_unit)
    #     return

    def ChangeBrTension(self, df_part):
        df_part_temp = df_part
        df_part_temp = df_part_temp[
            df_part_temp["condition"].str.contains('ｱﾝｸﾞﾙ')
            & df_part_temp["type"].str.contains('TS')]
        if len(df_part_temp.index.to_list()) > 0:
            for index in df_part_temp.index.to_list():
                print(index)
                df_part["type"][index] = 'Tr-B'
                df_part["unit"][index] = 'N/mm'
        return df_part

    def WriteData(self, df_input):
        print('writing data...')

        df = pd.read_excel(self.file_data, index_col=0)
        df_merge = pd.concat([df, df_input], axis=0, sort=False)
        df_merge.reset_index(inplace=True, drop=True)
        df_merge.to_excel(self.file_data, header=True, startcol=0)

        print(f'saved data file in {self.file_data}')

    def GetData(self, auto_file) -> (any):

        df = pd.read_excel(auto_file, header=None)

        df[2] = df[2].fillna('Press')
        df[2] = df[2] + df[3]

        for i in range(len(df)):
            row_num = 2 + i * 4
            if row_num < len(df):
                for j in range(1, 4):
                    df.at[row_num + 3, j] = df.at[row_num, j]

        df_data = pd.DataFrame()

        for i in range(len(df)):
            row_num = 5 + i * 4
            if row_num < len(df):
                df_data = df_data.append(df.loc[[row_num]], ignore_index=True)

        target_list_row = []
        for i, value in enumerate(df_data[1]):
            if self.target[:2] in str(value):
                target_list_row.append(i)

        df_data = df_data.loc[target_list_row]

        # select titles
        df_data = df_data[[1, 2, 9, 10, 11, 14, 15]]

        if self.test_mode:
            print(df_data)

        # return
        return df_data


def DoIt(target: str, test_mode=False):
    tension = Tension(target)
    tension.TestMode(mode=test_mode)
    try:
        tension.StartProcess()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = input('target name (ex: ABC001): ')
    # target = 'CBA001'
    DoIt(target, test_mode=True)
