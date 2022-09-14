import pathlib
import glob

filePath = fr'C:\Users\junsa\Desktop\junsai\*\**\*.x*'
list = glob.glob(filePath, recursive=True)

p_temp = pathlib.Path(r'C:\Users\junsa\Desktop\junsai').glob('**\*.x*')
print(p_temp)
for p in p_temp:
    print(p.name)