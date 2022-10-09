import time
import CollectFiles
import CollectAutoTension
import Compression
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




def DoProcess(user:str, user_family:str,target: str, target_dir_path:str):
    CollectFiles.Check(user, target, target_dir_path)
    CollectAutoTension.DoIt(target, user_family)
    Muni.DoIt(target)
    Rheometer.DoIt(target)
    AutoTension.DoIt(target)
    Hardness.DoIt(target)
    HeatResist.DoIt(target)
    OilTension.DoIt(target)
    Compression.DoIt(target)
    
    Treatment.DoIt(target)
    print('\n===================')
    print('   process done   ')
    print('===================\n')


if __name__ == '__main__':

    class Color:
        YELLOW = '\033[33m'
        RESET  = '\033[0m'
        BLUE   = '\033[34m'
        BG_BLUE     = '\033[44m'

    
    
    print(f'{Color.BLUE}=============================================================={Color.RESET}')
    print(f'{Color.BLUE} Hello!! this is Auto Handling Data System made by K.J.  ver 0.2 \n ')
    display_time = time.localtime() 
    print(f' Current time is {display_time.tm_year}/ {display_time.tm_mon}/ {display_time.tm_mday}   {display_time.tm_hour}:{display_time.tm_min}  \n')
    # for company
    target_dir_path = r'\\kfs03a\labo\9101-NVH_DATA\ホース'
    # target_dir_path = r'C:\Users\junsa\Desktop'
    print(f' path : {target_dir_path} {Color.RESET}')

    print(f'{Color.BLUE}=============================================================={Color.RESET}')

    print('python version: ' , platform.python_version())
    print('pandas version: '  , pd.__version__)

    is_ok_family_name = False
    if_ok_first_name = False

    # checking family name
    while not is_ok_family_name:
        user_family = input(' input your family name (ex: 田中) : ')

        # admin special
        if user_family == 'qq':
            user_family = '小暮'
            user_first = '準才'
            is_ok_family_name = True
            if_ok_first_name = True


        # too short or too long of the name
        elif len(user_family) == 0 or len(user_family) > 10:
            print(' something wrong try again')
        
        # maybe OK
        else:
            is_ok_family_name = True


    # checking first name
    while not if_ok_first_name:
        user_first = input('\n ※ if you want EPDM data, please input EPDM \n input your first name (ex: 花子) :  ')
        
        # too short or too long of the name
        if len(user_first) == 0 or len(user_first) > 10:
            print(' somthing wrong try again')

        # maybe OK
        else:
            if_ok_first_name =True

    user = user_family + user_first

    # for exception
    if user_first == "EPDM":
        user = 'EPDM'
    else:
        pass

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
                time.sleep(1)
                DoProcess(user, user_family, target, target_dir_path)
            else:
                pass
