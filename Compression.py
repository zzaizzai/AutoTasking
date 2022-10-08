import pandas as pd
import Service
import glob


target = input('target: ')
exp_name = '圧縮永久歪み_自動集積'
file_path = Service.data_dir(target) + rf'\{exp_name}*{target}*.xls*'

file_list = glob.glob(file_path)

print(file_list)


df = pd.read_excel(file_list[0], header=9)
print(len(df))
print(df.iloc[:,[1,7]])
df_data = df.iloc[:,[1,7]]


title= ['配合番号', '歪率']
df_data.columns = title
df_data['配合番号'] = df_data['配合番号'].ffill()
print(df_data)


target_list = df_data['配合番号'].to_list()

# remove duplication
target_list_set = set(target_list)
target_list = list(target_list_set)

print(target_list)

# find mean data
mean_data_index = []
for i in range(len(target_list)):
    mean_index = 3 + 4*i
    print(mean_index, df_data['配合番号'][mean_index])
    mean_data_index.append(mean_index)

print(mean_data_index)
df_data = df_data.loc[mean_data_index]


df_data = df_data.round(0)


df_data =  df_data.transpose()
df_data.columns = df_data.loc['配合番号']
df_data.drop(index=['配合番号'], inplace=True)
print(df_data)


unit = ['%']
method  = ['compression']
condition = ['normal']
type = ['distortion']
df_data.insert(0, 'unit', unit )
df_data.insert(0, 'type', type )
df_data.insert(0, 'condition', condition )
df_data.insert(0, 'method', method )
df_data.reset_index(inplace= True, drop= True)
print(df_data)


class Compression:
    def __init__(self, target):
        self.target = target

    def FindFile(self):
        print('find file')
        

def hizumi(target: str):
    zumizumi = Compression(target)



if __name__ == '__name__':
    target = input('target: ')
    hizumi(target)
    