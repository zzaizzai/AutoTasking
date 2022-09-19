
import glob
import os
import shutil
import string
from turtle import shearfactor
import openpyxl



class CheckFiles:


    def __init__(self, user: str, target: str, destinationFolderPath:str, DesktopPath:str):
        self.user= user
        self.target = target
        self.filePath = destinationFolderPath +  fr'\{user}\*\**\*.x*'
        self.destination = DesktopPath + fr'\{target} Data'
        self.fileNamePath = destinationFolderPath + fr'\{user}'


    def FindFiles(self):    
        os.makedirs(self.destination, exist_ok=True)
        list = glob.glob(self.filePath, recursive=True)
        
        file_copy_num: int = 0
        for file in list:
            if self.target in file:
                # print(file[len(self.filePath) -10:])
                self.Copyfiles(file)
                file_copy_num = file_copy_num + 1
        print(f'we found {file_copy_num} files!!')
        print()

    def Copyfiles(self, targetFile: string):
    # check what experiment you want
        # print(self.fileNamePath)
        experiments = os.listdir(self.fileNamePath)

        for experiment in experiments:
            if experiment in targetFile:
                # os.makedirs(self.destination + fr'\{experiment}', exist_ok=True)
                # shutil.copy2(targetFile, self.destination + fr'\{experiment}')
                print(targetFile)
                shutil.copy2(targetFile, self.destination + fr'\{experiment} {os.path.basename(targetFile)}')


    def MakeDataExcel(self):
        print('making data excel')
        wb = openpyxl.Workbook()
        sheet = wb.active
        wb.save( self.destination +  fr'\{self.target} Data.xlsx')

        


def Check(user: str, target: str, destinationFolderPath: str, DesktopPath: str):
    ff = CheckFiles(user, target, destinationFolderPath, DesktopPath)
    ff.FindFiles()
    ff.MakeDataExcel()

if __name__ == 'CheckFiles.py':
    Check()
