# utf-8
# Author: <NAME>
# Date: 2021-01-08
# this is an example for drawing the combo chart
import os
import pandas as pd
import matplotlib.pyplot as plt
# get the current path
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# get the excel file path
file_path = os.path.join(current_path, "sales_total1.xlsx")
print('file_path',file_path)
# read the excel file
pandas_data=pd.read_excel(file_path)
print(pandas_data)
# draw the combo chart
figure=plt.figure()
x=pandas_data['月份']
y1=pandas_data['销售额']
y2=pandas_data['利润']
plt.bar(x,y1,color='green')
plt.plot(x,y2,color='red')
plt.title('sales profit by month')
plt.show()

