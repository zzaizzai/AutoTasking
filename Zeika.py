import pandas as pd
import Service
import os
import glob


class Zeika:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target):

        self.target = target

        self.exp_name = '脆化'
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'
        self.file_now = ''

    def FindFile(self):
        print('find files...')
        print(self.file_path)

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)

        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} files(s)')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name}')
            return

    def ReadFile(self):
        print('read file...')
        print(self.file_now)

        df = pd.read_excel(self.file_now, index_col=None, header=None)
        # print(df_target)
        df = df.loc[:, :20]
        # print(df)

        def find_standard_index(df):
            index_standard = {"row": 0, "col": 0}

            is_Search = True
            for col_index in df:
                for row_index, cell_value in enumerate(df[col_index]):
                    # print(cell_value)
                    if cell_value == "脆化点" and col_index < 10:
                        print(row_index, col_index)
                        index_standard["row"] = row_index
                        index_standard["col"] = col_index - 1
                        break
            return index_standard

        index_standard = find_standard_index(df)
        # print(index_standard)

        df_Tb = df.iloc[index_standard["row"]:, index_standard["col"]:]
        df_Tb.reset_index(inplace=True, drop=True)
        titles_new = df_Tb.columns.to_list()
        titles_new[0] = '配合番号'
        titles_new[1] = '脆化点'
        df_Tb.columns = titles_new

        target_name_list = df_Tb["配合番号"].to_list()
        target_name_list_with_new_target = []
        # print(target_name_list)
        for value in target_name_list:
            if str(value) != 'nan':
                # print(value)
                target_name_list_with_new_target.append(
                    Service.target_number_as(int(value), self.target))
            else:
                # print(value)
                target_name_list_with_new_target.append(value)
        df_Tb["配合番号"] = target_name_list_with_new_target

        # print(df_Tb)



        # remove not witten number
        serial_number_list = df_Tb["配合番号"].to_list()

        coun_numbers = 0
        for i, name in enumerate(serial_number_list):
            if str(name) != 'nan':
                coun_numbers += 1
        # print(coun_numbers)
        df_Tb = df_Tb.loc[:coun_numbers, :]
        # print(df_Tb)

        def find_Tb_temperature(df_Tb):
            # print(df_Tb)
            df_Tb["ムハカイ"] = 99
            df_Tb["ゼンハカイ"] = 99
            for col_index in df_Tb:
                # print(col_index)
                for row_index, cell_value in enumerate(df_Tb[col_index]):
                    if str(cell_value) != 'nan' and len(str(cell_value)) != 6:
                        if str(cell_value) in ['0', '0.0']:
                            df_Tb.at[row_index,
                                     "ムハカイ"] = df_Tb.at[0, col_index]
                        if str(cell_value) in ['5', '5.0']:
                            df_Tb.at[row_index,
                                     "ゼンハカイ"] = df_Tb.at[0, col_index]

            df_Tb = df_Tb.loc[:, ['配合番号', 'ムハカイ', '脆化点', 'ゼンハカイ']]
            df_Tb = df_Tb.drop(index=[0])
            df_Tb.reset_index(inplace=True, drop=True)

            return df_Tb

        df_Tb = find_Tb_temperature(df_Tb)

        def trans_axis_and_change_titles(df_Tb):
            # print(df_Tb)

            df_Tb = df_Tb.transpose()
            df_Tb.columns = df_Tb.loc['配合番号', :].to_list()
            df_Tb = df_Tb.drop(index=['配合番号'])

            unit_list = ['℃']*len(df_Tb)
            type_list = df_Tb.index.to_list()
            condition_list = ['none']*len(df_Tb)

            df_Tb.insert(0, 'unit', unit_list)
            df_Tb.insert(0, 'type', type_list)
            df_Tb.insert(0, 'condition', condition_list)
            df_Tb.insert(0, 'method', Service.file_name_without_target(
                self.file_now, self.target))

            df_Tb.reset_index(inplace=True, drop=True)

            return df_Tb

        df_Tb = trans_axis_and_change_titles(df_Tb)
        self.WriteDate(df_Tb)

    def WriteDate(self, df_input):
        print('writing data...')

        # print(df_input)

        file_data = Service.data_dir(
            self.target) + fr'\{self.target} Data.xlsx'
        is_file = os.path.isfile(file_data)

        if is_file:
            pass
        else:
            print('no data file')
            # or you can make a data file
            return

        df_data = pd.read_excel(file_data, index_col=0)

        df_merge = pd.concat([df_data, df_input], sort=False)
        df_merge.reset_index(inplace=True, drop=True)

        try:
            # df_merge.to_excel(file_data, index=True, header=True)
            Service.save_to_data_excel(file_data, df_input, self.exp_name)
        except Exception as e:
            print(e)
        # print(df_merge)

        print(f'saved data file in {file_data}')


def DoIt(target: str, test_mode=False):
    zeizei = Zeika(target)
    zeizei.TestMode(mode=test_mode)
    try:
        zeizei.FindFile()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = input('target: ')
    DoIt(target)
