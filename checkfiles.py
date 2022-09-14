
import glob
import os
import shutil
import string


user = 'junsai'
target = 'CBA001'


# find excell files
filePath = fr'C:\Users\junsa\Desktop\{user}\*\**\*.x*'
destination = fr'C:\Users\junsa\Desktop\{target}'



experiments = ['Delta', 'Rheometer', 'Oil']


def Copyfiles(targetFile: string):
    # check what experiment you want
    for experiment in experiments:
        if experiment in targetFile:
            if 'xlsm' in targetFile:
                shutil.copy2(targetFile, destination + fr'\{experiment} {target}.xlsm')
            elif 'xlsx' in targetFile:
                shutil.copy2(targetFile, destination + fr'\{experiment} {target}.xlsx')




os.makedirs(destination, exist_ok=True)
list = glob.glob(filePath, recursive=True)
for file in list:
    print(file[len(filePath) -10:])
    if target in file:
        Copyfiles(file)


