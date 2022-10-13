import pandas as pd
import glob
import shutil
import os
import Service
from multiprocessing.dummy import Pool as ThreadPool

class hippari:

    def __init__(self, target: str, file_auto_list: list, user_family_name:str ):
        self.file_auto_list = file_auto_list
        self.target = target
        self.data_dir = Service.data_dir(target)
        self.user_family_name = user_family_name
    
    def CheckDeaktopPath(self):
        print('checking dir folder exist or not')
        print(self.data_dir)
        is_file = os.path.isdir(self.data_dir)
        if is_file:
            print('data directory exists')
        else:
            print('you have no data dir')
            os.makedirs(self.data_dir, exist_ok=True)


    def ReadFile(self):

        files_list = []

        def getData(file_auto: str):
            print(file_auto)
            is_file = os.path.isfile(file_auto)
            if is_file:
                pass
            else:
                print('no file')
                return
                
            df = pd.read_excel(file_auto)

            target_rows = df.query(f'配合番号 == "{self.target}"')
            # print(target_rows)
            # print(target_rows['ファイル名'])
            file_date_list = []
            for date in target_rows['測定日']:
                file_date_list.append(date)

            print(file_date_list)
            
            myname_rows = df.query(f'依頼者名 == "{self.user_family_name}"')
            for date in file_date_list:
                target_files = myname_rows.query(f'測定日 == "{date}"')
                # print('target_files')
                # print(target_files)
                for file_name in target_files['ファイル名']:
                    files_list.append(file_name)

        if len(self.file_auto_list) == 0:
            print("no file list of auto tension")
            return
        else:
            pass

        for file_auto in self.file_auto_list:
            getData(file_auto)

        # remove duplication
        files_list_set = set(files_list)
        files_list = list(files_list_set)

        print('target files')
        print(files_list)


        print('collecting wait a minute.......')
        
        self.CopyFiles(files_list)
        


    
    def CopyFiles(self, files_list):
        dir_auto_tension = self.data_dir + r'\auto_tension'
        os.makedirs(dir_auto_tension, exist_ok=True)

        def get_file_path(file):
            file_path = ''
            year = file[0:4]
            month = int(file[4:6])
            day = file[6:8]
            
            while len(file_path) < 1:
                files_path =  glob.glob(rf'\\kfs04\share2\4501-R_AND_D\JSK\全自動引張り\データ\2022年*\{month}*\{file[0:8]}\**\{file}*', recursive=True)
                print(files_path)
                file_path = files_path[0]
            return file_path

        # multi Thread
        pool = ThreadPool(4)
        
        file_path_list = list(pool.map(get_file_path ,files_list))
        # print(file_path_list)
        for file_path in file_path_list:
            shutil.copy2(file_path, dir_auto_tension)
        print('copy done')

def DoIt(target: str, user_family_name:str ):

    file_auto_list =  [r'\\kfs04\share2\4501-R_AND_D\JSK\全自動引張り\全自動引張り試験共通リスト2022年.xlsx']       


    hip = hippari(target, file_auto_list, user_family_name)

    try:
        hip.CheckDeaktopPath()
        hip.ReadFile() 
    except  Exception as e:
        print(e)



if __name__ == '__main__':
    user_family_name = input('user name (only family name): ')
    target = input('target Series ( ex: ABC001 ): ')
    
    # user_family_name = '小暮'
    # target = 'FJX001'
    
    DoIt(target, user_family_name)