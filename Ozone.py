import Service
import glob
import pandas as pd
import os


class Ozone:

    test_mode = False

    def TestMode(self, mode: bool):
        self.TestMode = mode

    def __init__(self, target: str):
        self.exp_name = 'オゾン '

        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_data = Service.data_dir(
            target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''

        self.exist_steam = True
    def StartProcess(self):
        print('Start Process')

        file_list = glob.glob(self.file_path)
        file_list  = sorted(file_list, key=len)
        print(file_list)

        if len(file_list) > 0 :
            print(f'found {len(file_list)} file(s)')
            pass
        else:
            print(f'No {self.exp_name}')
            return
        

        for file in file_list:
            self.file_now = file
            self.ReadData()
    
    def ReadData(self):
        print('reading files')
        print(self.file_now)

        sheet_list = pd.ExcelFile(self.file_now).sheet_names

        if '設定シート' in sheet_list:
            sheet_list.remove('設定シート')

        if 'きれつ表' in sheet_list:
            sheet_list.remove('きれつ表')


        df_all = pd.DataFrame()

        for sheet in sheet_list:
            try:
                df_all = pd.concat([df_all, self.ReadSheet(sheet)])
            except Exception as e:
                print(e)

        df_all = df_all.reset_index(drop=True)
        # print(df_all)

        self.WriteData(df_all)

    def ReadSheet(self, sheet):
        print('reading sheet', sheet)


        df_sheet = pd.read_excel(self.file_now, sheet_name=sheet, header=10, index_col=1)

        df_condition = df_sheet.iloc[:, 1]
        # print(df_condition)
        
        index_steam_start = 0
        for i, value in enumerate(df_condition.to_list()):
            # print(i ,value)
            if str(value) == 'スチーム':
                index_steam_start = i
                break

        if index_steam_start == 0:
            print('no steam samples')
            index_steam_start = len(df_condition.to_list())
            self.exist_steam = False
        else:
            print('index stream start', index_steam_start)
            
    
        df_sheet = df_sheet.iloc[:, 3:]
        del_list = []
        for _, col_name in enumerate(df_sheet.columns.to_list()):
            if "Unnamed" in str(col_name):
                del_list.append(str(col_name))

        if len(del_list) > 0:
            df_sheet = df_sheet.drop(del_list, axis=1)

        target_list = df_sheet.index.to_list()
        for i, value in enumerate(target_list):
            if str(value) == 'nan' and i > 0:
                target_list[i] = target_list[i-1]

        df_sheet.index = target_list
        if self.test_mode:
            print(df_sheet)


        df_press = df_sheet.iloc[:index_steam_start, :]    
        if self.test_mode:
            print('df_press')
            print(df_press)
        df_steam = pd.DataFrame()
        if index_steam_start != 0:
            print(index_steam_start)
            df_steam = df_sheet.iloc[index_steam_start:, :]

        if self.test_mode:
            print('df_steam')
            print(df_steam)


        df_press_n1 = df_press[~df_press.index.duplicated(keep='first')]
        df_press_n2 =  df_press[~df_press.index.duplicated(keep='last')]

        df_steam_n1 = df_steam[~df_steam.index.duplicated(keep='first')]
        df_steam_n2 = df_steam[~df_steam.index.duplicated(keep='last')]

        df_press_n1 = df_press_n1.transpose()
        df_press_n2 = df_press_n2.transpose()

        df_steam_n1 = df_steam_n1.transpose()
        df_steam_n2 = df_steam_n2.transpose()


        def add_units_press(df, number: int, steam=False):
            unit_list = df.index.to_list()
            # print(unit_list)
            df.insert(0, 'unit', unit_list)

            type_list = [f'n={number}']*len(df)
            df.insert(0, 'type', type_list)

            if steam:
                condition_list = [f'{sheet} スチーム']*len(df)
            else:
                condition_list = [f'{sheet} プレス']*len(df)

            df.insert(0, 'condition', condition_list)
            
            method_list = [self.exp_name]*len(df)
            df.insert(0, 'method', method_list)
            return df

        df_press_n1 = add_units_press(df_press_n1, 1)
        df_press_n2 = add_units_press(df_press_n2, 2)

        if self.exist_steam:
            df_steam_n1 = add_units_press(df_steam_n1, 1, steam=True)
            df_steam_n2 = add_units_press(df_steam_n2, 2, steam=True)

        if self.test_mode:
            print(df_press_n1)
            print(df_press_n2)
            print(df_steam_n1)
            print(df_steam_n2)


        df_input = pd.DataFrame()

        try:
            df_input = pd.concat([df_press_n1, df_press_n2])
            if self.exist_steam:
                df_input = pd.concat([df_input, df_steam_n1])
                df_input = pd.concat([df_input, df_steam_n2])
        except Exception as e:
            print(e)
        if self.test_mode:
            print(df_input)
        return df_input

    def WriteData(self, df_input):
        print('wrting df')

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
            df_merge.to_excel(file_data, index=True, header=True)
        except Exception as e:
            print(e)
        # print(df_merge)

        print(f'saved data file in {file_data}')        
        


def DoIt(target:str, test_mode =False):
    zozoni = Ozone(target)
    zozoni.TestMode(test_mode)
    try:
        zozoni.StartProcess()
    except Exception as e:
        print(e)


if __name__ == "__main__":
    target = input('target: ')
    DoIt(target, test_mode=True)

