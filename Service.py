
def target_number(number: int, target:str) -> (str):
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
