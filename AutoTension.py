import pandas as pd
import os
import glob


class Tension:


    def __init__(self, target: str ):
        self.DesktopPath = os.path.expanduser('~/Desktop')

        self.file_dir = self.DesktopPath  + rf'\{target} Data'
        # self.auto_file_dir = auto_file_dir
        self.auto_file_dir = self.file_dir + r'\auto_tension'
        self.target = target
        self.file_data = self.file_dir + fr'\{self.target} Data.xlsx'
        self.len_df_col = 0

    def GetFiles(self):

        is_file = os.path.isfile(self.file_data)


        if is_file:
            pass
        else:
            print('no data file')
            return

        auto_file_path = self.auto_file_dir + r'\*.xlsm'
        print(auto_file_path)
        auto_file_list = glob.glob(auto_file_path)
        print(auto_file_list)

        if len(auto_file_list) == 0 :
            print('no auto tesnion data file')
            return

        df_all =pd.DataFrame()

        for auto_file in auto_file_list:
            print(auto_file)
            df_all = pd.concat([df_all, self.GetData(auto_file)])
        
        print(df_all)
        
        print('showing merge df with sorted')
        # sort witt tile
        df_all = df_all.sort_values(by=[2, 1])

        df_all.reset_index(inplace= True, drop= True)
        df_all = df_all.T.reset_index(drop=True).T
        print(df_all)

        # remove another targets with checking target number and data file lengh
        df_len = pd.read_excel(self.file_data, index_col=0)
        self.len_df_col = len(df_len.columns) - 4 # without condition, method, type, unit
        print(f'len of col with only data : {self.len_df_col}')
        
        if self.len_df_col < 1:
            print('you should do another method first')
            return

        remove_row_list = []
        for i, value in enumerate(df_all[0]):
            print(int(value[3:])) 
            if int(value[3:]) < int(self.target[3:]):
                print('below range ')
                remove_row_list.append(i)
            if int(value[3:]) >=  self.len_df_col + int(self.target[3:]):
                print('over range')
                remove_row_list.append(i)
        print(remove_row_list)
        df_all = df_all.drop(remove_row_list)
        print(df_all)


        df_all.reset_index(inplace= True, drop= True)
        df_all = df_all.T.reset_index(drop=True).T
        

        target_condition_list = []
        for value in df_all[1]:
            target_condition_list.append(value)

        target_condition_set = set(target_condition_list)
        target_condition_list = list(target_condition_set)

        print(target_condition_list)

        target_condition_list = sorted(target_condition_list, key=len)
        for condition in target_condition_list:
            df_part = pd.DataFrame()

            for i, value in enumerate(df_all[1]):
                if value == condition:
                    df_part = df_part.append(df_all.loc[[i]])
            print(df_part)

            df_part = df_part.transpose()
            df_part = df_part.T.reset_index(drop=True).T
            df_part.reset_index(inplace= True, drop= True)
            print(df_part)

            unit = ['MPa', 'MPa', 'MPa' ,'MPa', '%']
            method = ['25%M', '50%M', '100%M','tension','elongation']
            condition_name = df_part[0][1]
            condition = [condition_name] * 5
            name = ['auto tesntion'] * 5
            df_part = df_part.loc[[2,3,4,5,6]]

            df_part.insert(loc=0, column='unit', value=unit)
            df_part.insert(loc=0, column='type', value=method)
            df_part.insert(loc=0, column='condition', value=condition)
            df_part.insert(loc=0, column='method', value=name)

            df_part = df_part.T.reset_index(drop=True).T
            df_part.reset_index(inplace= True, drop= True)

            print(df_part)

            self.WriteData(df_part)


        print('merge done')



    def WriteData(self, df_input):
        print('writing data...')

        



        df = pd.read_excel(self.file_data, index_col=0)

        df_merge = pd.concat([df, df_input])
        print(df_merge)

        df_merge.reset_index(inplace= True, drop= True)

        df_merge.to_excel(self.file_data, index=True, header=True, startcol=0)

        print(df_merge)
        print(f'saved data file in {self.file_data}')


    def GetData(self, auto_file)-> (any):

        df = pd.read_excel(auto_file , header=None)
        print(df)

        df[2] = df[2].fillna('Normal')
        df[2] = df[2] + df[3]

        for i in range(len(df)):
            row_num = 2 + i*4
            if row_num < len(df):
                for j in range(1, 4):
                    df.at[row_num + 3, j] = df.at[row_num, j]
        print(df)

        df_data = pd.DataFrame()

        for i in range(len(df)):
            row_num = 5 + i*4
            if row_num < len(df):
                df_data = df_data.append(df.loc[[row_num]], ignore_index=True)


        print('..df only average values...')
        print(df_data)
        print(self.target[:2])
        target_list_row = []
        for i, value in enumerate(df_data[1]):
            print(i)
            print(value)
            if self.target[:2] in str(value):
                print(f'include {self.target[0:2]} at {value}')
                target_list_row.append(i)

        print(target_list_row)

        df_data = df_data.loc[target_list_row]
        print(df_data)

        # select titles
        df_data = df_data[[1,2,11,12,13,14,15]]

        print(df_data)

        return df_data

def DoIt(target: str):
    tension = Tension(target)
    tension.GetFiles()

if __name__ == '__main__':
    target = input('target name (ex: ABC001): ')
    # target = 'CBA001'
    DoIt(target)