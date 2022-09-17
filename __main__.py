from time import sleep
import CheckFiles
import Rheometer



if __name__ == '__main__':
    user = 'junsai'
    target = 'CBA001'

    # user = input(' what is your name: ')
    # target = input(' what is you target : ')

    print(f'Hello {user}!')
    print(f'target is {target}')
    sleep(1)
    CheckFiles.Check(user, target)
    Rheometer.Rheo(target)
