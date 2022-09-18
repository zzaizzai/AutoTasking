from time import sleep
import CheckFiles
import Delta
import Rheometer




if __name__ == '__main__':

    user = 'junsai'
    target = 'CBA001'

    targetFolderPath = r'C:\Users\junsa\Desktop'
    DesktopPath = r'C:\Users\junsa\Desktop'
    # user = input(' what is your name: ')
    # target = input(' what is you target : ')

    print(f'Hello {user}!')
    print(f'target is {target}')
    sleep(1)
    CheckFiles.Check(user, target, targetFolderPath, DesktopPath)
    Rheometer.Rheo(target, DesktopPath)
    Delta.Del(target, DesktopPath)
    
