
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile

data = pd.read_excel(r'C:\Users\ThaoDang\Documents\Matlab\Test.xlsx','Test')

f=data.columns
print(f)
#print(f[1])
print(data[f[2]])

for x in range(2,len(f)):
    for y in range(0,len(data[f[x]])):
        a=data[f[x]]
        if (a[y] == np.nan):
            l="empty"
            print(l)
        else:
            print(f[x])
    




