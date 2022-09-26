import time
import CollectFiles
import CollectAutoTension
import Rheometer
import Muni
import AutoTension




def DoProcess(user:str, user_family:str,target: str, target_dir_path:str):
    CollectFiles.Check(user, target, target_dir_path)
    CollectAutoTension.DoIt(target, user_family)
    Rheometer.DoIt(target)
    Muni.DoIt(target)
    AutoTension.DoIt(target)

    print('process done')


if __name__ == '__main__':


    print(' Hello!! this is Auto Handling Data System made by K.J. \n ')
    display_time = time.localtime()
    print(f' Current time is {display_time.tm_year}/ {display_time.tm_mon}/ {display_time.tm_mday}   {display_time.tm_hour}:{display_time.tm_min} \n')

    # for company

    target_dir_path = r'\\kfs03a\labo\9101-NVH_DATA\ホース'
    # target_dir_path = r'C:\Users\junsa\Desktop'
    
    print(f'path : {target_dir_path} \n')
    user_family = input(' input your family name (ex: 田中) : ')
    user_first = input('\n ※ if EPDM data please input EPDM \n input your first name (ex: 花子) :  ')
    
    user = user_family + user_first

    if user_first == "EPDM":
        user = 'EPDM'
    else:
        pass

    print(f'Hello {user}!')

    target = input(' what is you target (ex: ABC001) : ')

    print(f'target is {target}')
    time.sleep(1)
    DoProcess(user, user_family, target, target_dir_path)
