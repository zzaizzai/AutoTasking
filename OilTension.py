import glob
import Service
import pandas as pd


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
        print(df)
        print(df.iloc[:,[1,9,10]])

        df_data = df.iloc[:,[8,9,10]]

        # delete NaN
        df_data = df_data.query("EB == EB")

        print(df_data)


        #  handle df titles
        df_data = df_data.transpose()

        print(df_data.columns.to_list())

        title_list = df_data.columns.to_list()
        title_new_list = ['nan']
        count = 0
        target_index = 0
        for i in range(1, len(title_list)):
            count += 1
            title_new_list.append(Service.target_number(target_index, self.target))
            if count == 4:
                count = 0
                target_index += 1
        df_data.columns = title_new_list
        print(title_new_list)

        
        # find mean value
        mean_col_index = []
        for i in range(3, len(title_new_list)):
            print(title_new_list[i])
            if title_new_list[i] != title_new_list[i-1]:
                print('it is after mena data')
                mean_col_index.append(i-1)
        print('mean_col_index', mean_col_index)

        
        print(df_data.iloc[:,mean_col_index])
        df_data = df_data.iloc[:,mean_col_index]

        unit = ['Mpa', 'Mpa', '%']
        type = ['100%M', 'tension', 'elongation']
        condition = [sheet]*3
        method =['oil tension']*3

        df_data.insert(0,'unit',unit)
        df_data.insert(0,'type', type)
        df_data.insert(0, 'condition', condition )
        df_data.insert(0, 'method', method)

        df_data.reset_index(inplace= True, drop= True)

        print(df_data)
        
        return df_data
    def WriteData(self, df_input):
        print('writing data')

        df = pd.read_excel(self.file_data, index_col=0)
        print(df)
        df_merge = pd.concat([df, df_input], axis=0, sort=False)

        print(df_merge)

        df_merge.reset_index(inplace=True, drop=True)

        df_merge.to_excel(self.file_data, index=True, header=True, startcol=0)
        print(f'saved data file in {self.file_data}')

def DoIt(target:str):
    oiru = OilTension(target)
    oiru.FindFile()


if __name__ == '__main__':
    target = input('target: ')
    print('Oil Tension')
    DoIt(target)