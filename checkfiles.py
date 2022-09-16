
import glob
import os
import shutil
import string



class CheckFiles:
    def __init__(self, user: str, target: str):
        self.user= user
        self.target = target
        self.filePath = fr'C:\Users\junsa\Desktop\{user}\*\**\*.x*'
        self.destination = fr'C:\Users\junsa\Desktop\{target}'
        self.fileNamePath = fr'C:\Users\junsa\Desktop\{user}'

    def FindFiles(self):    
        os.makedirs(self.destination, exist_ok=True)
        list = glob.glob(self.filePath, recursive=True)
        for file in list:
            print(file[len(self.filePath) -10:])
            if self.target in file:
                self.Copyfiles(file)

    def Copyfiles(self, targetFile: string):
    # check what experiment you want
        experiments = os.listdir(self.fileNamePath)
        
        for experiment in experiments:
            if experiment in targetFile:
                os.makedirs(self.destination + fr'\{experiment}', exist_ok=True)
                shutil.copy2(targetFile, self.destination + fr'\{experiment}')

ff = CheckFiles('junsai', 'CBA001')
ff.FindFiles()

# find excell files
# filePath = fr'C:\Users\junsa\Desktop\{user}\*\**\*.x*'
# destination = fr'C:\Users\junsa\Desktop\{target}'


# fileNamePath = fr'C:\Users\junsa\Desktop\{user}'
# experiments = os.listdir(fileNamePath)


# def Copyfiles(targetFile: string):
#     # check what experiment you want
#     for experiment in experiments:
#         if experiment in targetFile:
#             os.makedirs(destination + fr'\{experiment}', exist_ok=True)
#             shutil.copy2(targetFile, destination + fr'\{experiment}')

# def FindFiles():    
#     os.makedirs(destination, exist_ok=True)
#     list = glob.glob(filePath, recursive=True)
#     for file in list:
#         print(file[len(filePath) -10:])
#         if target in file:
#             Copyfiles(file)




