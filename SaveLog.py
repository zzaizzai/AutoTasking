import datetime
import json

def Save_Data(name:str, target: str):
    data = {
        'name': name,
        'target': target,
        'date': str(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')),
         }
    f = './log.txt'
    json_file = open(f, mode='a')
    json.dump(data, json_file, indent=2, ensure_ascii=False)
    # f.write(data)
    json_file.close()


if __name__ == '__main__':
    user_name = input('input user name:')
    target = input('target: ')
    Save_Data(str(user_name), str(target))