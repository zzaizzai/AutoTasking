from time import sleep
import CollectFiles




if __name__ == '__main__':

    user = 'junsai'
    # target = 'CBA001'

    # for company

    # targetFolderPath = r'\\kfs03a\labo\9101-NVH_DATA\ホース'
    targetFolderPath = r'C:\Users\junsa\Desktop'

    user_family = input(' what is your family name: ')
    user_first = input(' what is your first name: ')
    
    user = user_family + user_first

    print(f'Hello {user}!')

    target = input(' what is you target : ')

    print(f'target is {target}')
    sleep(1)
    CollectFiles.Check(user, target, targetFolderPath)
