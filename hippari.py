import pandas as pd
import glob
import shutil
import os

class hippari:

    def __init__(self, target: str, filePath: str, DesktopPath: str, user_family_name:str ):
        self.filePath = filePath
        self.target = target
        self.DataFolder = DesktopPath + fr'\{target} Data'
        self.user_family_name = user_family_name
    
    def CheckDeaktopPath(self):
        print('checking dir folder exist or not')
        print(self.DataFolder)
        is_file = os.path.isdir(self.DataFolder)
        if is_file:
            print('data directory exists')
        else:
            print('you have no data dir')
            os.makedirs(self.DataFolder, exist_ok=True)


    def ReadFile(self):
        print(self.filePath)
        df = pd.read_excel(self.filePath)
        target_rows = df.query(f'配合番号 == "{self.target}"')
        print(target_rows)
        print(target_rows['ファイル名'])
        file_date_list = []
        for date in target_rows['測定日']:
            file_date_list.append(date)

        print(file_date_list)
        files_list = []
        myname_rows = df.query(f'依頼者名 == "{user_family_name}"')
        for date in file_date_list:
            target_files = myname_rows.query(f'測定日 == "{date}"')
            print('target_files')
            print(target_files)
            for file_name in target_files['ファイル名']:
                files_list.append(file_name)
        
        # remove duplication
        files_list_set = set(files_list)
        files_list = list(files_list_set)
        print(files_list)

        print('target files')
        self.CopyFiles(files_list)
    
    def CopyFiles(self, files_list):
        for file in files_list:
            file_path_list = glob.glob(rf'\\kfs04\share2\4501-R_AND_D\JSK\全自動引張り\データ\2022年7～12月\**\{file}*', recursive=True)
            print(f'copying... {file_path_list}')
            shutil.copy2(file_path_list[0], self.DataFolder)
        print('copy done')
        

# if __name__ == 'hippari.py':
DesktopPath = os.path.expanduser('~/Desktop')
# user_family_name = input('user name (only family name): ')
# target = input('target Series ( ex: ABC001 ): ')

user_family_name = '小暮'
# DesktopPath = r'C:\Users\1010020990\Desktop'
target = 'FJX001'


path =  r'\\kfs04\share2\4501-R_AND_D\JSK\全自動引張り\全自動引張り試験共通リスト2022年.xlsx'
hip = hippari(target, path, DesktopPath, user_family_name )
hip.CheckDeaktopPath()
hip.ReadFile()
