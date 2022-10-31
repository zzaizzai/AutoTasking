import time
import CollectFiles
import CollectAutoTension
import Compression
import Deruta
import OilTension
import Rheometer
import Muni
import AutoTension
import Treatment
import HeatResist
import Hardness
import Service
import platform
import pandas as pd
import os
import Osidasi
from dotenv import load_dotenv




def DoProcess(user: str, user_family: str, target: str, target_dir_path: str, test_mode=False):
    CollectFiles.Check(user, target, target_dir_path)
    CollectAutoTension.DoIt(target, user_family, test_mode=test_mode)
    Muni.DoIt(target)
    Rheometer.DoIt(target)
    AutoTension.DoIt(target, test_mode=test_mode)
    Hardness.DoIt(target, test_mode=test_mode)
    HeatResist.DoIt(target, test_mode=test_mode)
    OilTension.DoIt(target, test_mode=test_mode)
    Deruta.DoIt(target, test_mode=test_mode)
    Compression.DoIt(target, test_mode=test_mode)

    Osidasi.DoIt(target, testMode=test_mode)

    Treatment.DoIt(target)
    print('\n=========================================================')
    # print('   process done   ')
    print(Service.work_done)
    print('=========================================================\n')


if __name__ == '__main__':
    load_dotenv()
    print(f'=========================================================================================')
    print(f' Hello!! this is Auto Handling Data System made by K.J.  ver 0.2 \n ')
    display_time = time.localtime()
    print(
        f' Current time : {display_time.tm_year}/ {display_time.tm_mon}/ {display_time.tm_mday}   {display_time.tm_hour}:{display_time.tm_min}  \n')
    # for company
    target_dir_path = str(os.environ.get('HOSE_DIR'))
    # target_dir_path = r'C:\Users\junsa\Desktop'

    # print(Service.rakuraku_hose)

    print(f' save_dir : {os.path.expanduser("~/Desktop")}')
    print(f' path : {target_dir_path}')

    print(f'=========================================================================================')
    print('10/24 Oil tension fixed, added Heat Resist with 3s hardness.')
    print('some rounding issure.')
    print('python version: ', platform.python_version())
    print('pandas version: ', pd.__version__)

    is_ok_family_name = False
    if_ok_first_name = False
    test_mode = False

    # checking family name
    while not is_ok_family_name:
        print(f' test mode: {test_mode}')
        user_family = input(' input your family name (ex: 田中) : ')

        # admin special
        if user_family == 'qq':
            user_family = '小暮'
            user_first = '準才'
            is_ok_family_name = True
            if_ok_first_name = True

        elif user_family == 'qqq':
            user_family = '小暮'
            user_first = '準才'
            is_ok_family_name = True
            if_ok_first_name = True
            test_mode = True
            print('test mode on')

        elif Service.check_user_name(user_family):
            if user_family == 'test':
                test_mode = True
                print('test mode on')
                pass
            else:
                is_ok_family_name = True

        else:
            print()
            pass

    # checking first name
    while not if_ok_first_name:
        user_first = input(
            '\n ※ if you want EPDM data, please input EPDM \n input your first name (ex: 花子) :  ')

        if Service.check_user_name(user_first):
            if_ok_first_name = True
        else:
            print()
            pass

    user = user_family + user_first

    # for exception
    if user_first == "EPDM":
        user = '※※住化EPDM代替関連テーマ(FG〇～)※※'
    else:
        pass

    if user == '小暮準才':
        print('Hello my lord')
    else:
        print(f'Hello {user}!')

    while True:
        target = input(' what is your target (ex: ABC001) or say "bye" : ')
        if target == "bye":
            print("I miss you! :(")
            time.sleep(3)
            break
        else:
            if Service.check_target(target):
                print(f'target is {target}')
                print(f'test mode: {test_mode}')
                time.sleep(1)

                DoProcess(user, user_family, target,
                          target_dir_path, test_mode=test_mode)

            else:
                print()
                pass
