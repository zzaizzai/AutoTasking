from email import header
import pandas as pd
import glob
import os

class Tension:


    def __init__(self, target: str ,auto_file_dir: str):
        self.DesktopPath = os.path.expanduser('~/Desktop')

        self.file_dir = self.DesktopPath  + rf'\{target} Data'
        self.auto_file_dir = auto_file_dir
        self.target = target

    def GetFiles(self):
        auto_file_list = glob.glob(self.auto_file_dir + r'\auto_tension\*.xlsx')
        print(auto_file_list)

        df_all =pd.DataFrame()

        for auto_file in auto_file_list:
            print(auto_file)
            df_all = pd.concat([df_all, self.GetData(auto_file)])
        
        print('showing merge df with sorted')
        df_all = df_all.sort_values(by=[2, 1], na_position='first')

        df_all.reset_index(inplace= True, drop= True)
        df_all = df_all.T.reset_index(drop=True).T

        print(df_all)

        group_name_list = []
        for value in df_all[1]:
            group_name_list.append(value)

        group_name_set =set(group_name_list)
        group_name_list = list(group_name_set)

        print('conditions')
        print(group_name_list)

        num_samples = len(group_name_list)
        num_rows_sampels  = int(len(df_all)/ len(group_name_list))

        print(f'{num_samples} samples with {num_rows_sampels} rows')

        # divide df and write the data to data excel file
        for i in range(num_samples):
            df_part = df_all.iloc[ 0 +i*num_rows_sampels : num_rows_sampels + i*num_rows_sampels   , : ]

            df_part = df_part.transpose()
            print(df_part)
            # unit = ['M', 'M', 'min' ]
            # method = ['M1', 'Vm', 'T1']
            name = ['auto tesntion', 'auto tesntion', 'auto tesntion']
            # df_part.insert(1, 2, unit)
            # df_part.insert(1, 1, method)
            df_part = df_part.loc[[2,3,4,5]]
            print(df_part)
            # df_part.reset_index(inplace= True, drop= True)
            # df_part = df_part.T.reset_index(drop=True).T
            df_part.insert(0, 0, name)
            


            # self.WriteData(df_part)



    def WriteData(self, df_input):
        print('writing data...')

        file_data = self.file_dir + fr'\{self.target} Data.xlsx'

        is_file = os.path.isfile(file_data)

        if is_file:
            pass
        else:
            print('no data file')
            return

        df = pd.read_excel(file_data, index_col=0)
        print(df)

        df_merge = pd.concat([df, df_input])
        print(df_merge)

        df_merge.reset_index(inplace= True, drop= True)

        # df_merge.to_excel(file_data, index=True, header=True, startcol=0)
        # print(f'saved data file in {file_data}')


    def GetData(self, auto_file)-> (any):

        df = pd.read_excel(auto_file , header=None)
        print(df)

        # for i in range(len(df)):
        #     row_num = 2 + i*5
        #     if row_num < len(df):
        #         print(df[2][row_num])
        #         if pd.isnull(df[2][2+ i*5]):
        #             df.at[2+ i*5, 2] = 0
        
        for i in range(len(df)):
            row_num = 2 + i*5
            if row_num < len(df):
                for j in range(1, 4):
                    df.at[row_num + 4, j] = df.at[row_num, j]
        print(df)

        df_data = pd.DataFrame()

        for i in range(len(df)):
            row_num = 6 + i*5
            if row_num < len(df):
                df_data = df_data.append(df.loc[[row_num]], ignore_index=True)

        print(df_data)

        target_list_row = []
        for i, value in enumerate(df_data[1]):
            print(value)
            if target[:3] in value:
                print(f'include CBA at {value}')
                target_list_row.append(i)
        
        print(target_list_row)

        df_data = df_data.loc[target_list_row]
        print(df_data)


        df_data = df_data[[1,2,6,7,8,9]]

        print(df_data)

        return df_data

def TenTen(target: str, auto_file_dir: str ):
    tension = Tension(target, auto_file_dir)
    tension.GetFiles()

if __name__ == '__main__':
    auto_file_dir = r'C:\Users\junsa\Desktop\auto'
    target = 'CBA001'

    TenTen(target, auto_file_dir)