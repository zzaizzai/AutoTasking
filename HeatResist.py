import os
import glob
import pandas as pd
import Service

class HeatResist:

    def __init__(self, target: str):
        self.exp_name = '熱老化_自動集積 '

        self.target = target

        self.file_path = Service.data_dir(target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_now = ''
    
    def FindFile(self):
        print('find files...')
        print(self.file_path)

        file_list = glob.glob(self.file_path)
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
        print(sheet_list)

        if '設定シート' in sheet_list:
            sheet_list.remove('設定シート')
        else:
            pass

        print(sheet_list)


        df_all = pd.DataFrame()

        for sheet in sheet_list:
            df_all = pd.concat([df_all, self.ReadDataSheet(sheet)])
        
        print('all od input data')

        print(df_all)

        self.WriteData(df_all)

    def ReadDataSheet(self, sheet: str):
            
        print(sheet)
        df = pd.read_excel(self.file_now, sheet_name=sheet, header=9, index_col=1)
        df = df.transpose()
        df.reset_index(inplace= True, drop= True)
        print(df.index)


        print(df)


        ## title
        print(df.loc[[3]].values.tolist()[0])
        col_index = []
        for i, value in enumerate(df.loc[[4]].values.tolist()[0]):
            print(i, value)
            if not 'nan' in str(value):
                print('not nan')
                col_index.append(i)

        print(col_index)
        print(df.iloc[:,col_index])

        df = df.iloc[:,col_index]

        title_index =  df.columns.values.tolist()
        for i, value in enumerate(title_index):
            if 'nan' in str(value):
                print('nan',i)
                title_index[i] = title_index[i-1]
        print(title_index)
        df.columns = title_index
        print(df)



        
        mean_col_index = [0]
        row_mean_str = df.loc[[2]].values.tolist()[0]
        print(row_mean_str)
        for i, value in enumerate(row_mean_str):
            print(value)
            if '中央値' in str(value):
                print('mean str', i)
                mean_col_index.append(i)
        
        #  0s 
        zero_col_index = [0]
        row_zero = df.loc[[12]].values.tolist()[0]
        print(row_zero)
        
        for i in range(10):
            print(row_zero[1+ i*4])
            if str(row_zero[1+i*4]) != 'nan' and 1+i*4 + 3 < len(row_zero) :
                row_zero[1+i*4 + 3] = row_zero[1+i*4] 
            else:
                break
            
        print(row_zero)
        df.loc[[12]] = row_zero
        print(df)


        for i, value in enumerate(row_zero):
            print(value)
            if str(value) != 'nan' and str(value) != '０秒' and i < len(row_zero) -3:
                print('it is zero hardness', i)
                # row_zero[i+ 3] = row_zero[i]

        print(mean_col_index)

        df_input = df.iloc[:,mean_col_index]
        # df_input = df_input.dropna()
        print(df_input)

        # change unit title
        unit_list = df_input.columns.tolist()
        unit_list[0] = 'type'
        print(unit_list)
        df_input.columns = unit_list


        print('query')
        df_input = df_input.query("type in ['M 100%', '抗張力 ＭＰａ', '破断伸び％', '０秒']")
        print(df_input)


        # condition
        condition = [sheet]*4
        method = ['heat tension']*4
        unit = ['dd','dd','dd','dd']
        df_input.insert(1, 'unit', unit)
        df_input.insert(0, 'condition', condition)
        df_input.insert(0, 'method', method)


        return df_input


    def WriteData(self, df_input):
        print('merging data....')

        file_data = Service.data_dir(self.target) + fr'\{self.target} Data.xlsx'
        is_file = os.path.isfile(file_data)
        if is_file:
            pass
        else:
            print('no data file')
            # or you can make a data file
            return

        df = pd.read_excel(file_data, index_col=0)
        print(df)

        df_merge = pd.concat([df, df_input], sort=False)
        print(df_merge)

        df_merge.reset_index(inplace= True, drop= True)

        df_merge.to_excel(file_data, index=True, header=True, startcol=0)
        print(f'saved data file in {file_data}')

def DoIt(target:str):
    heat = HeatResist(target)
    heat.FindFile()

if __name__ == '__main__':
    # target = 'FJX001'
    target = input('target: ')
    
    DoIt(target)
