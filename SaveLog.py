import json
import datetime
import os


def Save_Data(name:str, target: str):
    data = {
        'name': name,
        'target': target,
        'date': str(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')),
         }
         
    log_json = []
    if os.path.exists('./log.json'):
        json_file = open('./log.json', 'r')
        log_json = json.load(json_file)
        json_file.close()

    log_json.append(data)
    json_write = open('./log.json', 'w')
    json.dump(log_json, json_write, indent=4,  separators=(',',': '))

    json_write.close()

if __name__ == '__main__':
    user_name = input('input user name:')
    target = input('target: ')
    Save_Data(str(user_name), str(target))