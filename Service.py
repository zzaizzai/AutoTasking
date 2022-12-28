import os
import pandas as pd
import numpy
import re


def target_number(number: int, target: str) -> str:
    """
    if you want 2 after target number of ABC001:
        target_number(2,target) -> ABC003
    """
    alphabet = target[0:3]
    num = int(target[3:])
    alphabet_num = alphabet + str('%03d' % (num + number))
    return alphabet_num


def target_number_as(number: int, target: str) -> str:
    """
    target_number_as(24, ABC020) -> ABC024
    """
    alphabet = target[0:3]
    alphabet_num = alphabet + str('%03d' % (number))
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



def round_by_check_eachone(df_round, demical: int, list_of_round: list):
    for i in range(len(df_round)):
        for j in range(len(df_round.columns)):
            # print(i, j, df_round.iat[i, j], type(df_round.iat[i, j]))
            if type(df_round.iat[i, j]) == float or type(df_round.iat[i, j]) == int or type(df_round.iat[i, j]) == numpy.float64:
                if df_round.iat[i, 2] in list_of_round:
                    df_round.iat[i, j] = round(
                        df_round.iat[i, j] + 0.00001, demical)
                else:
                    df_round.iat[i, j] = round(df_round.iat[i, j] + 0.00001, 2)
    return df_round

def file_name_without_target_and_expname_distin_underbar(file_path: str, target: str, ex_name: str):
    condition_name = ""

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    file_name_distin = re.split(' |_|-', file_name)
    for text in file_name_distin:
        if target in text or text in ex_name.split("_"):
            pass
        else:
            condition_name += text + " "
    if condition_name == "":
        condition_name = "none"
    # print("condition_name:  ",condition_name)
    return condition_name

def file_name_without_target_distin_underbar(file_path: str, target: str) -> str:
    condition_name = ""
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    file_name_distin = re.split('[_ ]', file_name)
    for text in file_name_distin:
        if target in text:
            pass
        else:
            condition_name += text + " "
    if condition_name == "":
        condition_name = "none"
    return condition_name


def file_name_without_target_and_expname(file_path: str, target: str, ex_name: str):
    condition_name = ""
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    for text in file_name.split():
        if target not in text and ex_name not in text:
            condition_name += text + " "
    if condition_name == "":
        condition_name = "none"
    return condition_name


def normal_round(df, round_num: int):
    """
    Use it in the Only Treatment
    """
    df.iloc[:, 4:] = df.iloc[:, 4:] + 0.00001
    df.iloc[:, 4:] = df.iloc[:, 4:].round(round_num)
    return df


def remove_dufulicant_in_list(duful_list: list):
    return list(set(duful_list))


def save_to_data_excel(file_data: str, df_input, file_name: str):
    """
    save df to data file as merged df
    """
    df = pd.read_excel(file_data, index_col=0)

    df_merge = pd.concat([df, df_input], sort=False)

    df_merge.reset_index(inplace=True, drop=True)

    try:
        df_merge.to_excel(file_data, index=True, header=True, startcol=0)
        print(f'{file_name} saved data file in {os.path.basename(file_data)}')
    except Exception as e:
        print(e)


def create_method_condition_type_unit(df_create_MCTU: pd.DataFrame, method_list:list, condition_list:list, type_list:list, unit_list:list) -> pd.DataFrame:
    df_create_MCTU.insert(0, 'unit', unit_list)
    df_create_MCTU.insert(0, 'type', type_list)
    df_create_MCTU.insert(0, 'condition', condition_list)
    df_create_MCTU.insert(0, 'method', method_list)
    return df_create_MCTU


rakuraku_hose = """
"""


work_done = """
░██╗░░░░░░░██╗░█████╗░██████╗░██╗░░██╗  ██████╗░░█████╗░███╗░░██╗███████╗
░██║░░██╗░░██║██╔══██╗██╔══██╗██║░██╔╝  ██╔══██╗██╔══██╗████╗░██║██╔════╝
░╚██╗████╗██╔╝██║░░██║██████╔╝█████═╝░  ██║░░██║██║░░██║██╔██╗██║█████╗░░
░░████╔═████║░██║░░██║██╔══██╗██╔═██╗░  ██║░░██║██║░░██║██║╚████║██╔══╝░░
░░╚██╔╝░╚██╔╝░╚█████╔╝██║░░██║██║░╚██╗  ██████╔╝╚█████╔╝██║░╚███║███████╗
░░░╚═╝░░░╚═╝░░░╚════╝░╚═╝░░╚═╝╚═╝░░╚═╝  ╚═════╝░░╚════╝░╚═╝░░╚══╝╚══════╝
"""
