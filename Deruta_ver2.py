import pandas as pd
import Service
import glob
import os

class Deruta2:

    def set_testmode(self, mode: bool):

        self.test_mode = mode

    def __init__(self, target):
        self.test_mode = False

        self.exp_name = 'ΔV_自動集積'

        self.target = target
        self.file_path = Service.data_dir(
            target) + rf'\{self.exp_name}*{target}*.xls*'
        self.file_data = Service.data_dir(
            target) + fr'\{self.target} Data.xlsx'
        self.file_now = ''

    def StartProcess(self):
        print(f'find files of {self.exp_name}')

        if self.test_mode:
            print(self.file_path)

        file_list = glob.glob(self.file_path)
        file_list = sorted(file_list, key=len)

        if len(file_list) == 0:
            self.file_path = Service.data_dir(
                self.target) + rf'\ΔV*{self.target}*.xls*'
            file_list = glob.glob(self.file_path)
            file_list = sorted(file_list, key=len)

        if len(file_list) == 0:
            self.file_path = Service.data_dir(
                self.target) + rf'\⊿Ｖ*{self.target}*.xls*'
            file_list = glob.glob(self.file_path)
            file_list = sorted(file_list, key=len)

        if self.test_mode:
            print(file_list)

        if len(file_list) > 0:
            print(f'found {len(file_list) } {self.exp_name} file(s)')

            for file in file_list:
                self.file_now = file
                self.ReadFile()
        else:
            print(f'No {self.exp_name}')
            return

    def ReadFile(self):
        print('read file ', os.path.basename(self.file_now))
        
        
        sheet_list = pd.ExcelFile(self.file_now).sheet_names
        
        
        df_deltaV_all = pd.DataFrame()
        for sheet in sheet_list:
            if sheet == sheet_list[0]:
                data_SG = self.GetSGData(sheet)
                df_deltaV_all = pd.concat([df_deltaV_all, data_SG])
            
            data  = self.ReadSheet(sheet)
            df_deltaV_all = pd.concat([df_deltaV_all, data])
        
        
        self.writedata(df_deltaV_all)
    def GetSGData(self, sheetName: str):
        pd_sheet = pd.read_excel(self.file_now, sheet_name=sheetName, header=None)
        pd_sheet = pd_sheet.loc[2:,0:11]
        df_data = pd_sheet.iloc[3:,0:11]
        
        df_SG = df_data.iloc[:, [1,5]]
        values_target_list = df_SG.iloc[:,0].values
        index_target_list = []
        for i, name in enumerate(df_SG.iloc[:,0].values):
            if str(name) != "nan":
                index_target_list.append(i)
        
        for i in index_target_list:
            values_target_list[i+1] = values_target_list[i ]
            values_target_list[i] = "nan"
            
        df_temp = pd.DataFrame(values_target_list, columns=['taret'])
        df_SG = df_SG.reset_index()
        df_SG["SG"] = df_SG[5]
        df_SG["target"]= df_temp
        
        index_delta_mean = []
        for i, value in enumerate(df_SG["target"]):
            if str(value) != "nan":
                index_delta_mean.append(i)           
        
        df_SG = df_SG.loc[index_delta_mean,["target", "SG"]]
        df_SG = df_SG.reset_index()
        
        for i, value in enumerate(df_SG["target"]):
            df_SG.loc[i, "target"] = Service.target_number(i, self.target)

        df_SG.drop(columns=["index"], inplace=True)
        df_SG = df_SG.transpose()
        
        values_target_list = df_SG.loc["target",:].values.tolist()
        df_SG.columns = values_target_list
        df_SG = df_SG.iloc[[1], :]
        df_SG.reset_index(drop=True, inplace=True)
        
        df_data_result = Service.create_method_condition_type_unit(
            df_SG,
            Service.file_name_without_target(self.file_now, self.target),
            "SG",
            "SG",
            'g/cm3'   
        )
        
        return df_data_result
    
    def ReadSheet(self, sheetName: str):
        
        pd_sheet = pd.read_excel(self.file_now, sheet_name=sheetName, header=None)
        pd_sheet = pd_sheet.loc[2:,0:11]
        
        df_information = pd_sheet.iloc[[0],0:11]
        df_data = pd_sheet.iloc[3:,0:11]


        # get liquid information 
        information_liquid = df_information.iat[0, 3]
        information_temperature = df_information.iat[0, 7]
        information_time = df_information.iat[0, 8]
        information_all = f"{information_liquid} {information_temperature}℃×{information_time}H"
        
    
        df_deltaV = df_data.iloc[:,[1,9]] 

        values_target_list = df_deltaV.iloc[:,0].values
        
        
        index_target_list = []
        for i, name in enumerate(df_deltaV.iloc[:,0].values):
            if str(name) != "nan":
                index_target_list.append(i)
        
        for i in index_target_list:
            values_target_list[i+1] = values_target_list[i ]
            values_target_list[i] = "nan"
        
        
        
        df_temp = pd.DataFrame(values_target_list, columns=['taret'])
        df_deltaV = df_deltaV.reset_index()
        df_deltaV["delta"] = df_deltaV[9]
        df_deltaV["target"]= df_temp
        
        index_delta_mean = []
        for i, value in enumerate(df_deltaV["target"]):
            if str(value) != "nan":
                index_delta_mean.append(i)
        
        df_deltaV = df_deltaV.loc[index_delta_mean,["target", "delta"]]
        df_deltaV = df_deltaV.reset_index()
        
        for i, value in enumerate(df_deltaV["target"]):
            df_deltaV.loc[i, "target"] = Service.target_number(i, self.target)

        df_deltaV.drop(columns=["index"], inplace=True)
        df_deltaV = df_deltaV.transpose()
        
    
        values_target_list = df_deltaV.loc["target",:].values.tolist()
        df_deltaV.columns = values_target_list
        df_deltaV = df_deltaV.iloc[[1], :]
        df_deltaV.reset_index(drop=True, inplace=True)

        df_data_result = Service.create_method_condition_type_unit(
            df_deltaV,
            Service.file_name_without_target(self.file_now, self.target),
            information_all,
            "⊿V",
            '%'   
        )
        
        return df_data_result
                
                
    def writedata(self, df_input):
        
        print('writing data')

        is_file = os.path.isfile(self.file_data)

        if is_file:
            pass
        else:
            print('no data file')
            return

        Service.save_to_data_excel(self.file_data, df_input, self.exp_name)


def DoIt(target: str, test_mode=False):
    ruta = Deruta2(target)
    ruta.set_testmode(mode=test_mode)

    try:
        ruta.StartProcess()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    target = input('target: ')

    DoIt(target, test_mode=True)
