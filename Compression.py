import pandas as pd
import Service
import glob


target = input('target')
exp_name = '圧縮永久歪み_自動集積'
file_path = Service.data_dir(target) + rf'\{exp_name}*{target}*.xls*'

file_list = glob.glob(file_path)

print(file_list)


df = pd.read_excel(file_list[0], header=9)
print(len(df))
print(df.iloc[:,[1,7]])
df_data = df.iloc[:,[1,7]]


title= ['配合番号', '硬度']
df_data.columns = title
df_data['配合番号'] = df_data['配合番号'].ffill()
print(df_data)


# print(df_data['配合番号'])
# hardness_value_mean = []
# for i in range(1, len(df_data['配合番号'])):
#     print(df_data['配合番号'][i])
#     if str(df_data['配合番号'][i]) != 'nan' and str(df_data['配合番号'][i-1]) == 'nan':
#         print('good')
#         hardness_value_mean.append(i-1)
# print(hardness_value_mean)


# print(df_data.loc[hardness_value_mean])

# df_data = df_data.dropna()