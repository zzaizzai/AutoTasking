import pandas as pd
import os
import glob
import win32com.client as win32


class Deruta:
    def __init__(self, target: str):
        self.DesktopPath = os.path.expanduser('~/Desktop')
        self.exp_name = '⊿Ｖ'

        self.target = target
        self.file_dir = self.DesktopPath + fr'\{target} Data'
        self.file_path_xls = self.file_dir + rf'\{self.exp_name} {target}*.xls'
        self.file_path_xlsx = ''
    
    def FindFile(self) -> bool:
        print('find file Delta....')
        file_list = glob.glob(self.file_path_xls)

        if len(file_list) > 0:
            print('xls file exist')
            self.file_path_xlsx = file_list[0]
            return True
        else:
            return False







# target = 'FJX001'
# deru = Deruta(target)
# if deru.FindFile():
#     print('file exist')
# else:
#     print('no file')

file_xls = r'C:\Users\1010020990\Desktop\FJX001 Data\⊿Ｖ FJX001-005.xls'
file_xlsm = r'C:\Users\1010020990\Desktop\FJX001 Data\⊿Ｖ FJX001-005.xlsm'
is_file = os.path.isfile(file_xlsm)
if is_file:
    print('file exist')
    pass
else: 
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(file_xls)
    wb.SaveAs(file_xlsm, FileFormat=51)
    wb.Close()
    excel.Application.Quit()
    print('made xlsx')
