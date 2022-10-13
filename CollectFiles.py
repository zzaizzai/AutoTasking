
import glob
import os
import shutil
import string
import openpyxl
import win32com.client as win32
import Service


class CollectFiles:


    def __init__(self, user: str, target: str, destination_dir_path:str):
        self.user= user
        self.target = target
        self.filePath = destination_dir_path +  fr'\{user}\*\**\*.x*'
        self.data_dir = Service.data_dir(target)
        self.fileNamePath = destination_dir_path + fr'\{user}'


    def FindFiles(self):    
        os.makedirs(self.data_dir, exist_ok=True)
        list = glob.glob(self.filePath, recursive=True)
        
        file_copy_num: int = 0
        file_copy_failed_num: int = 0
        for file in list:
            if self.target in file:
                try:
                    self.Copyfiles(file)
                except OSError as e:
                    print(e)
                    file_copy_failed_num += 1
                file_copy_num = file_copy_num + 1
        print(f'we found {file_copy_num} files!! with {file_copy_failed_num} failed')
        print()

    def Copyfiles(self, targetFile: string):
        experiments = os.listdir(self.fileNamePath)

        for experiment in experiments:
            if experiment in targetFile:

                print(targetFile)
                try:
                    shutil.copy2(targetFile, self.data_dir + fr'\{experiment} {os.path.basename(targetFile)}')
                except Exception as e:
                    print(e)
                    



    def MakeDataExcel(self):
        print('making data excel')
        wb = openpyxl.Workbook()
        wb.save( self.data_dir +  fr'\{self.target} Data.xlsx')

    def TranslateFromXlsToXlsx(self):
        print('translating xls files to xlsx file....')
        file_list = glob.glob(self.data_dir + r'\*.xls')
        print(file_list)

        for file_xls in file_list:
            is_file = os.path.isfile(file_xls + 'x')
            if is_file:
                print(f'xlsx is already exist {file_xls}')
            else:
                print(f'translate... {file_xls} + x')
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                wb = excel.Workbooks.Open(file_xls)
                excel.DisplayAlerts = False
                excel.Visible  = False
                wb.SaveAs(file_xls+"x", FileFormat = 51) #FileFormat = 51 is for .xlsx extension
                wb.Close()
                excel.Application.Quit()
                os.remove(file_xls)                



def Check(user: str, target: str, destination_dir_path: str):
    ff = CollectFiles(user, target, destination_dir_path)
    ff.FindFiles()
    ff.MakeDataExcel()
    ff.TranslateFromXlsToXlsx()

if __name__ == '__main__':
    # xl = EnsureDispatch("Word.Application")
    # print(sys.modules[xl.__module__].__file__)
    user = input('your fill name: ')
    target = input('target: ')
    # user = 'junsai'
    # target = 'CBA001'
    targetFolderPath = r'C:\Users\junsa\Desktop'
    
    Check(user, target, targetFolderPath)
