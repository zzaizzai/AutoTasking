import pandas as pd
from . import Service

class Deruta:

    def __init__(self, target):
        self.exp_name = '熱老化_自動集積 '
        
        self.target = target
        self.file_path = Service.data_dir(target) + rf'\{self.exp_name}*{target}*.xls*'