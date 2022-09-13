
import glob
import os
import shutil
import string

target = 'CBA001'


# find excell files
user = 'junsai'
filePath = fr'C:\Users\junsa\Desktop\{user}\*\**\*.xlsx'

destination = fr'C:\Users\junsa\Desktop\{target}'

os.makedirs(destination, exist_ok=True)

experiments = ['Delta', 'Rheometer']


def copyfiles(targetFile: string):

    # experiments you want
    experiments = ['Delta', 'Rheometer']

    for experiment in experiments:
        if experiment in targetFile:
            shutil.copyfile(targetFile, destination + fr'\{experiment} {target}.xlsx')



list = glob.glob(filePath, recursive=True)
for file in list:
    print(file[len(filePath) -11:])
    if target in file:
        copyfiles(file)


