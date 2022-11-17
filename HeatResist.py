import os
import glob
import pandas as pd
import Service


class HeatResist:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target: str):

        self.target = target

        self.exp_name = '熱老化_自動集積 '
        self.exp_name_2 = '自動引張り '

        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_path_2 = Service.data_dir(
            target) + rf'\{self.exp_name_2}*{target}*.xls*'

        self.file_now = ''

    def StartProcess(self):
        print('find files...')
        print(self.file_path)

        file_list = glob.glob(self.file_path) + glob.glob(self.file_path_2)
        file_list = sorted(file_list, key=len)
        print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list)} {self.exp_name} file(s)')

            for file in file_list:
                self.file_now = file
                self.ReadFile()

        else:
            print(f'No {self.exp_name}')
            return

    def ReadFile(self):
        print('read file...')
        print(self.file_now)
        sheet_list = pd.ExcelFile(self.file_now).sheet_names
        # print(sheet_list)

        if '設定シート' in sheet_list:
            sheet_list.remove('設定シート')
        else:
            pass

        # print(sheet_list)

        df_all = pd.DataFrame()

        for sheet in sheet_list:
            df_all = pd.concat([df_all, self.ReadDataSheet(sheet)])

        df_all = self.SortByTemperature(df_all)

        df_all = self.ChangeConditionName(df_all)

        self.WriteData(df_all)
    def ChangeConditionName(self, df_all):
        df_all = df_all.reset_index(drop=True)
        
        df_all_temp_condition = df_all["condition"]
        # print(df_all)
        for i, value in enumerate(df_all_temp_condition):
            # print(value.split("℃×"))
            temperature_value = value.split("℃×")[0]
            housr_value = value.split("℃×")[1].split("H")[0]
            # print(temperature_value, housr_value)
            df_all_temp_condition[i] = temperature_value +  "℃×" + housr_value + "H"
        df_all["condition"] = df_all_temp_condition
        # print(df_all)
        return df_all
        
    def SortByTemperature(self, df_all):
        # print(df_all)

        df_all = df_all.reset_index(drop=True)

        df_all_tem = df_all
        df_all_tem["temperature"] = df_all_tem["condition"]
        df_all_tem["hours"] = df_all_tem["condition"]

        for i, value in enumerate(df_all_tem["condition"]):
            condition_split = value.split("℃×")
            temperature =  condition_split[0]
            hours = condition_split[1].split("Hｒ")[0]

            df_all_tem["temperature"][i] = int(temperature)
            df_all_tem["hours"][i] = int(hours)
        # # df_all["temperature"] = df_all_temperature
        # # df_all["hour"] = df_all_hours
        df_all_tem.sort_values(by = ["temperature", "hours"], ascending=[True,True], inplace=True)
        df_all_tem.drop(columns=["temperature", "hours"], inplace = True)
        # print(df_all_tem)

        df_all = df_all_tem

        return df_all

    def ReadDataSheet(self, sheet: str):

        print('reading sheet: ', sheet)
        df = pd.read_excel(self.file_now, sheet_name=sheet,
                           header=9, index_col=1)
        df = df.transpose()
        df.reset_index(inplace=True, drop=True)
        # print(df.index)

        # print(df)

        # title
        col_index = []
        for i, value in enumerate(df.loc[[4]].values.tolist()[0]):
            # print(i, value)
            if not 'nan' in str(value):
                # print('not nan')
                col_index.append(i)


        df = df.iloc[:, col_index]

        title_index = df.columns.values.tolist()
        for i, value in enumerate(title_index):
            if 'nan' in str(value):
                # print('nan',i)
                title_index[i] = title_index[i-1]
        # print(title_index)
        df.columns = title_index
        # print(df)

        mean_col_index = [0]
        row_mean_str = df.loc[[2]].values.tolist()[0]
        # print(row_mean_str)
        for i, value in enumerate(row_mean_str):
            # print(value)
            if '中央値' in str(value):
                # print('mean str', i)
                mean_col_index.append(i)

        #  0s
        zero_col_index = [0]
        row_zero = df.loc[[12]].values.tolist()[0]
        row_three = df.loc[[13]].values.tolist()[0]

        for i in range(10):
            # print(row_zero[1+ i*4])
            if str(row_zero[1+i*4]) != 'nan' and 1+i*4 + 3 < len(row_zero):
                row_zero[1+i*4 + 3] = row_zero[1+i*4]
            else:
                break
        for i in range(10):
            # print(row_three[1+ i*4])
            if str(row_three[1+i*4]) != 'nan' and 1+i*4 + 3 < len(row_three):
                row_three[1+i*4 + 3] = row_three[1+i*4]
            else:
                break

        # print(row_zero)
        df.loc[[12]] = row_zero
        df.loc[[13]] = row_three
        # print(df)


        df_input = df.iloc[:, mean_col_index]

        # change unit title
        unit_list = df_input.columns.tolist()
        unit_list[0] = 'type'
        # print(unit_list)
        df_input.columns = unit_list

        df_input = df_input.query(
            "type in ['M 100%', '抗張力 ＭＰａ', '破断伸び％', '０秒','3秒']")
        df_input = df_input.iloc[:, 1:]
        # print('after query')
        # print(df_input)

        # condition
        condition = [sheet]*len(df_input)
        method = ['熱老化']*len(df_input)
        unit = ['MPa', 'MPa', '%', 'HA', 'HA']
        type_list = ['M100', 'TS', 'EB', 'HA(0s)', 'HA(3s)']
        df_input.insert(0, 'unit', unit)
        df_input.insert(0, 'type', type_list)
        df_input.insert(0, 'condition', condition)
        df_input.insert(0, 'method', method)

        # print(df_input)
        return df_input

    def WriteData(self, df_input):
        print('merging data....')

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
    heat = HeatResist(target)
    heat.TestMode(mode=test_mode)
    try:
        heat.StartProcess()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    # target = 'FJX001'
    target = input('target: ')

    DoIt(target)
