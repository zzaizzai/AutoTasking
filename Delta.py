import pandas as pd

class Delta:

    DesktopPath = r'C:\Users\junsa\Desktop'

    def __init__(self, target: str):
        self.destination = self.DesktopPath + fr'\{target} Data'
        self.filePath = self.destination + fr'\Delta {target}.xlsx'
    
    def ReadFile(self):
        df = pd.read_excel(self.filePath, header=None)
        # df = pd.read_excel(self.filePath, names=['硬度'])
        df = df.fillna(method='ffill')
        # data = df[[1,2,3]]
        data = df[[3]].loc[[5]]
        data = df[[3]].loc[[5]].values[0]
        print(data)




delta = Delta('CBA001')
delta.ReadFile()