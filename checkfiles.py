
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

def Check(user: str, target: str):
    ff = CheckFiles(user, target)
    ff.FindFiles()

if __name__ == 'CheckFiles.py':
    Check()




