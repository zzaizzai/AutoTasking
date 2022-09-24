from time import sleep
import CollectFiles




if __name__ == '__main__':

    print('Hello!! this is Auto Handling Data System \n ')

    # for company

    # targetFolderPath = r'\\kfs03a\labo\9101-NVH_DATA\ホース'
    targetFolderPath = r'C:\Users\junsa\Desktop'

    user_family = input(' input your family name (ex: 田中) : ')
    user_first = input(' @ if EPDM data please input EPDM \n input your first name (ex: 花子) :  ')
    
    user = user_family + user_first

    if user_first == "EPDM":
        user = 'EPDM'
    else:
        pass

    print(f'Hello {user}!')

    target = input(' what is you target (ex: ABC001) : ')

    print(f'target is {target}')
    sleep(1)
    CollectFiles.Check(user, target, targetFolderPath)
