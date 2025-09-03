# utf-8
# Author: <NAME>
# Date: 2021-09-13
# this is a pandas example
import pandas as pd
# 1d series
list1 = [1,2,3]
ser1 = pd.Series(list1)
print("ser1 is: ",ser1)
# 2d series
list2 = [[1,2,3],[4,5,6]]
ser2 = pd.Series(list2)
print("ser2 is: ",ser2)
