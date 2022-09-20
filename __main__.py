from time import sleep
import Checkfiles
import Delta
import Rheometer
import os




if __name__ == '__main__':

    # user = '小暮準才'
    # target = 'FJX001'

    user = 'junsai'
    target = 'CBA001'

    # for company
    DesktopPath = os.path.expanduser('~/Desktop')
    # targetFolderPath = r'\\kfs03a\labo\9101-NVH_DATA\ホース'
    
    targetFolderPath = r'C:\Users\junsa\Desktop'
    # DesktopPath = r'C:\Users\junsa\Desktop'
    # user = input(' what is your name: ')
    # target = input(' what is you target : ')

    print(f'Hello {user}!')
    print(f'target is {target}')
    sleep(1)
    Checkfiles.Check(user, target, targetFolderPath, DesktopPath)
    # Rheometer.Rheo(target, DesktopPath)
    # Delta.Del(target, DesktopPath)
