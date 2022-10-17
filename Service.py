import os
import pandas as pd


def target_number(number: int, target: str) -> str:
    """
    if you want 2 after target number of ABC001:
        target_number(2,target) -> ABC003
    """
    alphabet = target[0:3]
    num = int(target[3:])
    alphabet_num = alphabet + str('%03d' % (num + number))
    return alphabet_num


def check_target(target: str) -> bool:
    if len(target) != 6:
        print(' the target must be 6 characters like ABC123')
        return False
    elif target[:3].isalpha() == False or target[3:].isdigit() == False:
        print(' target rule wrong')
        return False
    else:
        return True


def check_user_name(user_name: str) -> bool:
    """
    check user name is good for the environment
    True: It is ok
    """
    if len(user_name) == 0 or len(user_name) > 8:
        print(' user name is too short or too long')
        return False
    elif user_name.isdigit():
        print(' input only string')
        return False
    else:
        return True


def data_dir(target: str) -> str:
    """
    Data file in your desktop
    """
    desktop = os.path.expanduser('~/Desktop')
    return desktop + rf'\{target} Data'


def file_name_without_target(file_path: str, target: str) -> str:
    result_name = ""
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    for text in file_name.split():
        if target not in text:
            result_name += text + " "
    if result_name == "":
        result_name = "none"
    return result_name


def file_name_without_target_and_expname(file_path: str, target: str, ex_name: str):
    condition_name = ""
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    for text in file_name.split():
        if target not in text and ex_name not in text:
            condition_name += text + " "
    if condition_name == "":
        condition_name = "none"
    return condition_name

def normal_round(df, round_num:int):
    df.iloc[:,4:] = df.iloc[:,4:] + 0.00001
    df.iloc[:,4:] = df.iloc[:,4:].round(round_num)
    return df


def save_to_data_excel(file_data: str, df_input):
    """
    save df to data file as merged df
    """
    df = pd.read_excel(file_data, index_col=0)
    # print(df)

    df_merge = pd.concat([df, df_input], sort=False)
    # print(df_merge)

    df_merge.reset_index(inplace=True, drop=True)

    df_merge.to_excel(file_data, index=True, header=True, startcol=0)

    print(f'saved data file in {file_data}')


rakuraku_hose = """
██╗░░██╗███████╗██╗░░░░░██╗░░░░░░█████╗░  ░██╗░░░░░░░██╗░█████╗░██████╗░██╗░░░░░██████╗░
██║░░██║██╔════╝██║░░░░░██║░░░░░██╔══██╗  ░██║░░██╗░░██║██╔══██╗██╔══██╗██║░░░░░██╔══██╗
███████║█████╗░░██║░░░░░██║░░░░░██║░░██║  ░╚██╗████╗██╔╝██║░░██║██████╔╝██║░░░░░██║░░██║
██╔══██║██╔══╝░░██║░░░░░██║░░░░░██║░░██║  ░░████╔═████║░██║░░██║██╔══██╗██║░░░░░██║░░██║
██║░░██║███████╗███████╗███████╗╚█████╔╝  ░░╚██╔╝░╚██╔╝░╚█████╔╝██║░░██║███████╗██████╔╝
╚═╝░░╚═╝╚══════╝╚══════╝╚══════╝░╚════╝░  ░░░╚═╝░░░╚═╝░░░╚════╝░╚═╝░░╚═╝╚══════╝╚═════╝░
"""


work_done = """
░██╗░░░░░░░██╗░█████╗░██████╗░██╗░░██╗  ██████╗░░█████╗░███╗░░██╗███████╗
░██║░░██╗░░██║██╔══██╗██╔══██╗██║░██╔╝  ██╔══██╗██╔══██╗████╗░██║██╔════╝
░╚██╗████╗██╔╝██║░░██║██████╔╝█████═╝░  ██║░░██║██║░░██║██╔██╗██║█████╗░░
░░████╔═████║░██║░░██║██╔══██╗██╔═██╗░  ██║░░██║██║░░██║██║╚████║██╔══╝░░
░░╚██╔╝░╚██╔╝░╚█████╔╝██║░░██║██║░╚██╗  ██████╔╝╚█████╔╝██║░╚███║███████╗
░░░╚═╝░░░╚═╝░░░╚════╝░╚═╝░░╚═╝╚═╝░░╚═╝  ╚═════╝░░╚════╝░╚═╝░░╚══╝╚══════╝
"""
