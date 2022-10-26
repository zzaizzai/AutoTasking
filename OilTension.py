import glob
import Service
import pandas as pd
import os 


class OilTension:

    def __init__(self, target:str):
        self.exp_name = '耐油引張り '

        self.target = target
        self.file_path = Service.data_dir(target) + rf'\{self.exp_name}*{target}*.xls*'

        self.file_data = Service.data_dir(target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''

    def FindFile(self):
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

        if '設定シート' in  sheet_list:
            sheet_list.remove('設定シート')
        else:
            pass

        print(sheet_list)

        df_all = pd.DataFrame()

        for sheet in sheet_list:
            df_all = pd.concat([df_all, self.ReadDataSheet(sheet)], sort=False)
        
        
        print('all df input data')
        print(df_all)

        self.WriteData(df_all)

        
    def ReadDataSheet(self, sheet: str):
        print('read data shett')
        df = pd.read_excel(self.file_now, sheet_name=sheet, header=9, index_col=1)
        # print(df)
        # print(df.iloc[:,[1,9,10, 11]])
        # df = df.iloc[:80,:]
        # df = df.iloc[:80,:]
        # print(df)
        df_data = df.iloc[:,[8,9,10, 11]]
        print('datadata', sheet)
        print(df_data)

        # set init target name for not written in first cell
        list_index = df_data.index.to_list()
        if str(list_index[1]) == 'nan':
            list_index[1] = str(self.target)
        
        df_data.index = list_index    
        print(df_data)
        

        # print(df_data[['JIS A']])

        # delete NaN
        df_data = df_data.query("EB == EB")

        # print(df_data)

        df_data.iloc[:,[-1]] = df_data.iloc[:,[-1]].ffill()
        

        print(df_data.index.to_list())

        index_list = df_data.index.to_list()
        index_values = Service.remove_dufulicant(index_list)
        print(index_values)

        number_target = len(index_values) - 1

        # set all target names
        print(index_list)
        for i in range(number_target):
            index_list[i*4] = Service.target_number(i, self.target)
            index_list[i*4 + 1] = Service.target_number(i, self.target)
            index_list[i*4 + 2] = Service.target_number(i, self.target)
            index_list[i*4 + 3] = Service.target_number(i, self.target)
        
        print(index_list)
        df_data.index = index_list
        print(df_data)
        


        # mean data
        index_mean = []
        for i in range(number_target):
            index_mean.append(3 + 4*i) 
        print(index_mean)

        df_mean = df_data.iloc[index_mean,:]
        print(df_mean)

        df_data = df_mean.transpose()
        # return

        # #  handle df titles
        # df_data = df_data.transpose()

        # # print(df_data.columns.to_list())

        # title_list = df_data.columns.to_list()
        # title_new_list = ['nan']
        # count = 0
        # target_index = 0
        # for i in range(1, len(title_list)):
        #     count += 1
        #     title_new_list.append(Service.target_number(target_index, self.target))
        #     if count == 4:
        #         count = 0
        #         target_index += 1
        # df_data.columns = title_new_list
        # # print(title_new_list)

        
        # # find mean value
        # mean_col_index = []
        # for i in range(3, len(title_new_list) -1):
        #     # print(title_new_list[i])
        #     # print
        #     if title_new_list[i] != title_new_list[i-1]:
        #         # print('it is after mena data')
        #         mean_col_index.append(i-1)
        # print('mean_col_index', mean_col_index)

        print(df_data)

        # return
        # print(df_data.iloc[:,mean_col_index])
        # df_data = df_data.iloc[:,mean_col_index]

        unit = ['MPa', 'MPa', '%','HA']
        type_list = ['100%M', 'TS', 'EB', 'HA(0s)']
        condition = [sheet]*4
        method =[Service.file_name_without_target(self.file_now, self.target)]*4

        df_data.insert(0,'unit',unit)
        df_data.insert(0,'type', type_list)
        df_data.insert(0, 'condition', condition )
        df_data.insert(0, 'method', method)

        df_data.reset_index(inplace= True, drop= True)

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

def DoIt(target:str):
    oiru = OilTension(target)
    try:
        oiru.FindFile()
    except Exception as e:
        print(e)



if __name__ == '__main__':
    target = input('target: ')
    print('Oil Tension')
    DoIt(target)
