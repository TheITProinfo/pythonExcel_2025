import os
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
file_path = os.path.join(current_path, "sales_total2.xlsx")
print('file_path',file_path)
padas_data=pd.read_excel(file_path)
print(padas_data)
# xial label
x=padas_data['月份']
y=padas_data['销售额']
plt.bar(x,y,color='green')



# draw the chart title
plt.title('sales turnover by month')
# drwa the x label
plt.xlabel('month',labelpad=10)
# draw the y label
plt.ylabel('sales turnover',labelpad=20 )
# draw the grid
plt.grid(axis='y',color='black',linestyle='--',linewidth=0.5)
# draw the legend
plt.legend(loc="upper right",fontsize=30,shadow=True)
# draw the chart
plt.show()
# draw the chart







