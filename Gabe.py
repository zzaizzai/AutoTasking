import glob
import Service
import pandas as pd
import os


class GabeGabe:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target: str):
        self.exp_name = 'ガーベ記録 '

        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_data = Service.data_dir(
            target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''

    def StartProcess(self):
        print('Find files')

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)

        # print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s) ')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name} ')
            return

    def ReadFile(self):


        df = pd.read_excel(self.file_now,)

        df = df.drop(columns=["Unnamed: 0"])
        method = ["ガーベイダイ"]*2
        condition_list = ["BK"]*2
        type_list = ["Surface", 'Edge']
        unit_list = ["Eval"]*2
        df = Service.create_method_condition_type_unit(df,method , condition_list, type_list, unit_list)



        self.WriteData(df)

    def WriteData(self, df_input):
        print('writing data')

        if os.path.isfile(self.file_data):
            pass
        else:
            print('no data file')
            return

        Service.save_to_data_excel(self.file_data, df_input, self.exp_name)


def DoIt(target: str, test_mode=False):
    oiru = GabeGabe(target)
    oiru.TestMode(mode=test_mode)
    try:
        oiru.StartProcess()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = input('target: ')
    print('Gabedai')
    DoIt(target, test_mode=True)
