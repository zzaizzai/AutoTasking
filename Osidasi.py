import pandas as pd
import Service
import glob
import os


class Osidasi:

    testMode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target: str):
        self.exp_name = '押出し '
        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_data = Service.data_dir(
            target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''

    def StartProcess(self):
        print('find file')
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
        print('Read file')

        df = pd.read_excel(self.file_now, header=16, index_col=1)
        # print(df)
        # print(df.index.to_list())

        target_list = df.index.to_list()
        # print(len(target_list))
        index_target_init = []
        for i, value in enumerate(target_list):
            # print(i , value)
            if str(value) != 'nan' and len(value) == 6:
                index_target_init.append(i)
        # print(index_target_init)

        for index in index_target_init:
            target_list[index-1] = target_list[index]
            target_list[index+1] = target_list[index]
            target_list[index+2] = target_list[index]

        df.index = target_list
        # print(df)

        target_list_numbers = Service.remove_dufulicant(target_list)

        # without nan
        number_of_target: int = len(target_list_numbers) - 1
        # print('number_of_target',number_of_target)

        df = df.iloc[:number_of_target*4+1, :]
        # print(df)

        # print(df.columns.to_list())
        titles = df.columns.to_list()
        titles[0] = 'mean'
        df.columns = titles
        # print(len(df))

        # mean data
        # print(target_list)
        count_for_mean = 0
        index_mean = []
        index_eval = []
        index_tempa = []
        for i in range(len(target_list)):
            # print(target_list[i])
            if str(target_list[i]) != 'nan':
                count_for_mean += 1
                if count_for_mean % 4 == 0:
                    # print('mean')
                    index_mean.append(i)
                    index_eval.append(i-2)
                    index_tempa.append(i-3)
                if count_for_mean > number_of_target*4:
                    break

        # get only mean data
        df_mean = df.iloc[index_mean, :]
        df_mean = df_mean.loc[:, ['L', 'W', 'Swell', 'Swell.1']]
        # print(df_mean)

        # get evaluations data
        df_eval = df.iloc[index_eval, :]
        df_eval = df_eval.loc[:, ['H.1', 'Sc', 'R.F']]
        # print(df_eval)

        # get temp data
        df_tempa = df.iloc[index_tempa, :]
        df_tempa = df_tempa.loc[:, ['D', 'R']]
        # print(df_tempa)

        # get Sulfurization data
        df_sulf = df.iloc[index_mean, :]
        df_sulf = df_sulf.loc[:, ['加硫前', '加硫後']]
        # print(df_sulf)

        # print(df.loc[:,'C.H'].to_list()[1:])
        values_pressure_CH = df.loc[:, 'C.H'].to_list()[1:]
        values_pressure_H = df.loc[:, 'H'].to_list()[1:]

        count_pressure = 0
        list_pressure_H = []
        for i, value in enumerate(values_pressure_H):
            # print(i, value)
            if str(value) != 'nan':
                count_pressure += 1
                if count_pressure % 2 == 0:
                    list_pressure_H.append(
                        str(values_pressure_H[i-1]) + ' → ' + str(values_pressure_H[i]))
        # print(list_pressure_H)

        count_pressure_CH = 0
        list_pressure_CH = []
        for i, value in enumerate(values_pressure_CH):
            # print(i, value)
            if str(value) != 'nan':
                count_pressure_CH += 1
                if count_pressure_CH % 2 == 0:
                    list_pressure_CH.append(
                        str(values_pressure_CH[i-1]) + ' → ' + str(values_pressure_CH[i]))
        # print(list_pressure_CH)

        df_pressrue = df.iloc[index_eval, :]
        df_pressrue = df_pressrue.loc[:, ['H', 'C.H']]

        df_pressrue.loc[:, 'H'] = list_pressure_H
        df_pressrue.loc[:, 'C.H'] = list_pressure_CH
        # print(df_pressrue)

        df_all = pd.DataFrame()

        df_mean = df_mean.transpose()
        df_eval = df_eval.transpose()
        df_tempa = df_tempa.transpose()
        df_pressrue = df_pressrue.transpose()
        df_sulf = df_sulf.transpose()

        df_all = pd.concat([df_all, df_eval])
        df_all = pd.concat([df_all, df_tempa])
        df_all = pd.concat([df_all, df_pressrue])
        df_all = pd.concat([df_all, df_mean])
        df_all = pd.concat([df_all, df_sulf])

        # print(df_all.loc[['L','W','Swell.1'],:])

        # print(df_all)

        self.Treat(df_all)

    def Treat(self, df_all: pd.DataFrame):

        condition = ['none']*len(df_all)
        method = ['osidashi(beta)']*len(df_all)
        unit = ['ss']*len(df_all)
        type_list = df_all.index.to_list()

        type_list[type_list.index('Swell')] = 'swell dia.'
        type_list[type_list.index('Swell.1')] = 'swell thick.'

        df_all.insert(0, 'unit', unit)
        df_all.insert(0, 'type', type_list)
        df_all.insert(0, 'condition', condition)
        df_all.insert(0, 'method', method)

        df_all.reset_index(inplace=True, drop=True)

        df_all = Service.round_by_check_eachone(df_all, 0, ['L'])

        print(df_all)

        self.Write(df_all)

    def Write(self, df_input):
        print('writing')

        file_data = Service.data_dir(
            self.target) + fr'\{self.target} Data.xlsx'
        is_file = os.path.isfile(file_data)

        if is_file:
            pass
        else:
            print('no data file')
            # or you can make a data file
            return

        Service.save_to_data_excel(file_data, df_input)


def DoIt(target: str, testMode=False):
    osiosi = Osidasi(target)
    osiosi.TestMode(testMode)
    try:
        osiosi.StartProcess()
    except Exception as e:
        print(e)


if __name__ == "__main__":
    target = input('target: ')
    DoIt(target, testMode=True)
