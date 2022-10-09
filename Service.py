import os
from typing import Any
import pandas as pd


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


def save_to_data_excel(file_data:str, df_input: Any):
    """
    save df to data file as merged df
    """
    df = pd.read_excel(file_data, index_col=0)
    print(df)

    df_merge = pd.concat([df, df_input], sort=False)
    print(df_merge)

    df_merge.reset_index(inplace=True, drop=True)

    df_merge.to_excel(file_data, index=True, header=True, startcol=0)

    print(f'saved data file in {file_data}')

