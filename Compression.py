import pandas as pd
import Service
import glob
import os


class Compression:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target):
        self.target = target
        self.exp_name = '圧縮永久歪み_自動集積'
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}* {target}*.xls*'
        self.file_now = ''

    def FindFile(self):
        print('find file')
        print(self.file_path)

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)
        print(file_list)
        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s)')
            pass
        else:
            print(f'No {self.exp_name}')
            print('search as 圧縮永久歪')
            self.file_path = Service.data_dir(
                self.target) + rf'\圧縮永久歪* {self.target}*.xls*'
            file_list = glob.glob(self.file_path)
            file_list = sorted(file_list, key=len)
            if len(file_list) > 0:
                print(f'found {len(file_list)} {self.exp_name} file(s)')
                pass
            else:
                print(f'No {self.exp_name}')
                return

        for file in file_list:
            print(file)
            self.file_now = file
            self.ReadFile()

    def ReadFile(self):
        print('read File ', os.path.basename(self.file_now))

        sheet_list = pd.ExcelFile(self.file_now).sheet_names
        print(sheet_list)

        df_all = pd.DataFrame()

        for sheet in sheet_list:
            df_all = pd.concat([df_all, self.ReadDataSheet(sheet)])

        df_all = self.SortByTemperature(df_all)


        df_all = pd.concat([df_all, self.get_hardness(sheet_list[0])])

        self.WriteData(df_all)

    def get_hardness(self, sheet_name_the_first:str):
        print("get hardness 0-3")
        print(sheet_name_the_first)
        df = pd.read_excel(self.file_now, sheet_name=sheet_name_the_first, header=9)
        df_hardness = df.loc[:,["Unnamed: 1","瞬間値","3秒値"]]

        print()

        number_of_target: int = 0
        for i in range(len(df_hardness)):
            if str(df_hardness.iat[i,0]) != "nan" and i % 4 == 0 :
                number_of_target += 1
                df_hardness.iat[i+1,0]= df_hardness.iat[i,0]
                df_hardness.iat[i+2,0]= df_hardness.iat[i,0]
                df_hardness.iat[i+3,0]= df_hardness.iat[i,0]

        index_row_middle = [4*i + 3 for i in range(number_of_target)]

        df_middle_value = df_hardness.loc[index_row_middle ,:]

        

        # rounding
        df_middle_value["瞬間値"]  = df_middle_value["瞬間値"] + 0.001
        df_middle_value["瞬間値"]  = df_middle_value["瞬間値"].round(0)
        df_middle_value["瞬間値"]  = df_middle_value["瞬間値"].apply(int).apply(str)

        df_middle_value["H0-H3"] = df_middle_value["瞬間値"] 

        df_middle_value["3秒値"]  = df_middle_value["3秒値"] + 0.001
        df_middle_value["3秒値"]  = df_middle_value["3秒値"].round(0)
        df_middle_value["3秒値"]  = df_middle_value["3秒値"].apply(int).apply(str)
        
        # H0-H3
        df_middle_value["H0-H3"] = df_middle_value["H0-H3"].str.cat(df_middle_value["3秒値"], sep=" - ")
        df_middle_value = df_middle_value.loc[:,["Unnamed: 1", "H0-H3"]]
        df_middle_value = df_middle_value.transpose()
        df_middle_value.reset_index(inplace=True, drop=True)

        df_middle_value.columns = df_middle_value.iloc[0]
        df_middle_value = df_middle_value.drop(df_middle_value.index[0])

        unit = ["HA"]
        type_list = ["H0-H3"]
        condition = ["BK"]
        method = [Service.file_name_without_target(self.file_now, self.target)]

        df_middle_value.insert(0, 'unit', unit)
        df_middle_value.insert(0, 'type', type_list)
        df_middle_value.insert(0, 'condition', condition)
        df_middle_value.insert(0, 'method', method)

        return df_middle_value





    def SortByTemperature(self, df_all):

        df_all = df_all.reset_index(drop=True)
        df_all_tem = df_all
        df_all_tem["temperature"] = df_all_tem["condition"]
        df_all_tem["hours"] = df_all_tem["condition"]

        for i, value in enumerate(df_all_tem["condition"]):
            condition_split = value.split("℃×")
            df_all_tem["temperature"][i] = int(value.split("℃×")[0])
            df_all_tem["hours"][i] = int(value.split("℃×")[1].split("H")[0])
        # print(df_all_tem)
        df_all_tem.sort_values(by=["temperature", "hours"], ascending=[
                               True, True], inplace=True)
        df_all_tem.drop(columns=["temperature", "hours"], inplace=True)

        return df_all_tem

    def ReadDataSheet(self, sheet: str):
        df = pd.read_excel(self.file_now, sheet_name=sheet, header=9)

        # find data col index
        col_index_data : int = 7
        for i in range(10):
            if  str(df.iat[0,i]) == "9.38":
                col_index_data = i + 1
                break

        df = df.iloc[:, [1, col_index_data]]

        title = ['配合番号', '歪率']
        df.columns = title
        df['配合番号'] = df['配合番号'].ffill()

        target_list = df['配合番号'].to_list()

        target_list_set = set(target_list)
        target_list = list(target_list_set)

        # find mean data

        mean_data_index = []
        for i in range(len(target_list)):
            mean_index = 3 + 4*i
            mean_data_index.append(mean_index)

        df = df.loc[mean_data_index]

        # rounding to int
        df = df.round(0)

        df = df.transpose()
        df.columns = df.loc['配合番号']
        df.drop(index=['配合番号'], inplace=True)

        unit = ['%']
        # method = ['compression']
        method = [Service.file_name_without_target(self.file_now, self.target)]
        condition = [sheet]
        type_list = ['CS']

        df.insert(0, 'unit', unit)
        df.insert(0, 'type', type_list)
        df.insert(0, 'condition', condition)
        df.insert(0, 'method', method)
        df.reset_index(inplace=True, drop=True)

        return df

    def WriteData(self, df_input):
        print('merging data...')

        file_data = Service.data_dir(
            self.target) + fr'\{self.target} Data.xlsx'
        is_file = os.path.isfile(file_data)

        if is_file:
            pass
        else:
            print('no data file')
            return

        Service.save_to_data_excel(file_data, df_input, self.exp_name)


def DoIt(target: str, test_mode=False):
    zumizumi = Compression(target)
    zumizumi.TestMode(test_mode)
    try:
        zumizumi.FindFile()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = input('target: ')
    DoIt(target, test_mode=True)
