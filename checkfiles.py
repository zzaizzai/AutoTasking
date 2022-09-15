
import glob
import os
import shutil
import string


user = 'junsai'
target = 'CBA001'


# find excell files
filePath = fr'C:\Users\junsa\Desktop\{user}\*\**\*.x*'
destination = fr'C:\Users\junsa\Desktop\{target}'


fileNamePath = fr'C:\Users\junsa\Desktop\{user}'
experiments = os.listdir(fileNamePath)


def Copyfiles(targetFile: string):
    # check what experiment you want
    for experiment in experiments:
        if experiment in targetFile:
            os.makedirs(destination + fr'\{experiment}', exist_ok=True)
            shutil.copy2(targetFile, destination + fr'\{experiment}')

def FindFiles():    
    os.makedirs(destination, exist_ok=True)
    list = glob.glob(filePath, recursive=True)
    for file in list:
        print(file[len(filePath) -10:])
        if target in file:
            Copyfiles(file)


FindFiles()


