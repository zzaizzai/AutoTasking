from time import sleep
import CheckFiles
import Rheometer



if __name__ == '__main__':
    user = 'junsai'
    target = 'CBA001'

    print(f'Hello {user}!')
    print(f'target is {target}')
    sleep(1)
    CheckFiles.Check(user, target)
    Rheometer.Rheo(target)
