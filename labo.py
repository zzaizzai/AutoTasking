from cmath import exp
import pandas as pd


def hello():
    df = pd.read_excel("",index_col=0)
    print("ddd")

try:
    hello()
except Exception as e:
    print(e)