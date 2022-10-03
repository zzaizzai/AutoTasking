import os


def target_number(number: int, target:str) -> str:
    """
    if you want 2 after target number of ABC001:
        target_number(2,target) -> ABC003
    """
    alphabet = target[0:3]
    num = int(target[3:])
    alphabet_num = alphabet + str('%03d' % (num + number))
    return alphabet_num

def check_target(target:str) -> bool:
    if len(target) != 6:
        print('the target must be 6 characters like ABC123')
        return False
    elif target[:3].isalpha() == False or target[3:].isdigit() == False :
        print('target rule wrong')
        return False
    else:
        return True

desktop = os.path.expanduser('~/Desktop')

def data_dir(target:str) -> str:
    """
    Data file in your desktop
    """
    return desktop  + rf'\{target} Data'

