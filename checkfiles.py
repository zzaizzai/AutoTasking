
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
        for file in list:
            if self.target in file:
                # print(file[len(self.filePath) -10:])
                self.Copyfiles(file)

    def Copyfiles(self, targetFile: string):
    # check what experiment you want
        print(self.fileNamePath)
        experiments = os.listdir(self.fileNamePath)

        for experiment in experiments:
            if experiment in targetFile:
                # os.makedirs(self.destination + fr'\{experiment}', exist_ok=True)
                # shutil.copy2(targetFile, self.destination + fr'\{experiment}')

                shutil.copy2(targetFile, self.destination + fr'\{experiment} {os.path.basename(targetFile)}')
    
    def MakeDataExcel(self):
        print('making data excel')
        wb = openpyxl.Workbook()
        sheet = wb.active
        for i in range(8):
            sheet.cell(row=1, column=5 + i,  value=f'{self.target}({i+1})')
        wb.save( self.destination +  fr'\{self.target} Data.xlsx')

        


def Check(user: str, target: str, destinationFolderPath: str, DesktopPath: str):
    ff = CheckFiles(user, target, destinationFolderPath, DesktopPath)
    ff.FindFiles()
    ff.MakeDataExcel()

if __name__ == 'CheckFiles.py':
    Check()
