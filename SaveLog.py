import csv
import os
import datetime


def Save_Data(name: str, target: str):
    if os.path.isfile('./log.csv'):
        pass
    else:
        print('new log file')
        f_new = open('./log.csv', 'w')
        f_new.write('id, name, target, time')
        f_new.close()

    f = open('./log.csv', 'r')
    reader = csv.DictReader(f)
    
    count_oldest = 0
    for row in reader:
        if int(row["id"]) >= count_oldest:
            count_oldest = int(row["id"])
    f.close()

    f = open('./log.csv', 'a')
    count_new = count_oldest + 1
    new_data = f"\n{count_new}, {name}, {target}, {datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')}"
    f.write(new_data)
    f.close()


if __name__ == "__main__":
    user_name = input("input user name: ")
    target = input("target: ")
    Save_Data(user_name, target)
