import glob
import Service
import pandas as pd
import os


class OilTension:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target: str):
        self.exp_name = '耐油引張り '

        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_data = Service.data_dir(
            target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''

    def StartProcess(self):
        print('Find file')
        print(self.file_path)

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)

        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s) ')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name} ')
            return

    def ReadFile(self):
        print('ReadFile')
        print('read file...')

        sheet_list = pd.ExcelFile(self.file_now).sheet_names
        print(sheet_list)

        if '設定シート' in sheet_list:
            sheet_list.remove('設定シート')
        else:
            pass

        print(sheet_list)

        df_all = pd.DataFrame()

        for sheet in sheet_list:
            df_all = pd.concat([df_all, self.ReadDataSheet(sheet)], sort=False)

        # print('all df input data')
        # print(df_all)

        self.WriteData(df_all)

    def ReadDataSheet(self, sheet: str):
        print(f'reading data of {sheet}')
        df = pd.read_excel(self.file_now, sheet_name=sheet,
                           header=9, index_col=1)

        df_data = df.iloc[:, [8, 9, 10, 11]]
        # df_data = df.iloc[:,[8,9,10]]
        print(sheet)
        # print(df_data)

        # set init target name for not written in first cell
        list_index = df_data.index.to_list()
        if str(list_index[1]) == 'nan':
            list_index[1] = str(self.target)

        df_data.index = list_index
        # print(df_data)

        # print('index list')
        # print(df.index.to_list())

        index_list = df_data.index.to_list()
        index_values = Service.remove_dufulicant(index_list)
        # print('removed deful')
        # print(index_values)

        number_target = len(index_values) - 1

        # set all target names, sometimes there is no name in target cells.... becuz of mistakes..?
        for i in range(number_target):
            index_for_target = i*4 + 1
            index_list[index_for_target] = Service.target_number(
                i, self.target)
            index_list[index_for_target +
                       1] = Service.target_number(i, self.target)
            index_list[index_for_target +
                       2] = Service.target_number(i, self.target)
            index_list[index_for_target +
                       3] = Service.target_number(i, self.target)

        df_data.index = index_list
        # print(df_data)

        # mean data
        index_mean = []
        for i in range(number_target):
            index_mean.append(3 + 4*i + 1)
        print(index_mean)
        df_mean = df_data.iloc[index_mean, :]
        print('mean data')

        # get hardness data
        index_mean_hardness = list(map(lambda x: x-3, index_mean))
        df_hardness = df_data.iloc[index_mean_hardness, :]
        df_mean.loc[:, 'HS'] = df_hardness.loc[:, 'HS']

        df_data = df_mean.transpose()

        unit = ['MPa', 'MPa', '%', 'HA']
        type_list = ['100%M', 'TS', 'EB', 'HA(0s)']
        condition = [sheet]*4
        method = [Service.file_name_without_target(
            self.file_now, self.target)]*4

        df_data.insert(0, 'unit', unit)
        df_data.insert(0, 'type', type_list)
        df_data.insert(0, 'condition', condition)
        df_data.insert(0, 'method', method)

        df_data.reset_index(inplace=True, drop=True)

        print(df_data)

        return df_data

    def WriteData(self, df_input):
        print('writing data')

        is_file = os.path.isfile(self.file_data)

        if is_file:
            pass
        else:
            print('no data file')
            return

        Service.save_to_data_excel(self.file_data, df_input)


def DoIt(target: str, test_mode=False):
    oiru = OilTension(target)
    oiru.TestMode(mode=test_mode)
    try:
        oiru.StartProcess()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = input('target: ')
    print('Oil Tension')
    DoIt(target, test_mode=True)
