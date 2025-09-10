# utf8
# author: <NAME>
# date: 2021-01-04
# this is an example read data from excel then generating bar chart
import xlwings as xw
import pandas as pd
import os
import matplotlib.pyplot as plt
# get the current path
current_path = os.path.dirname(os.path.abspath(__file__))

print("current file path: ",current_path)

file_path = os.path.join(current_path, "sales_total.xlsx")
print('file_path',file_path)
padas_data=pd.read_excel(file_path)
print(padas_data)

figure=plt.figure()
x=padas_data['月份']
y=padas_data['销售额']
plt.bar(x,y,color='green')
plt.title('sales profit by month')
plt.show()

# call the excel
app = xw.App(visible=True,add_book=False)
# open the workbook
workbook = app.books.open(file_path)
# get all sheets in the workbook
sheets = workbook.sheets
# add new sheet to the workbook
sheet = workbook.sheets.add('bar chart9900', after=sheets[-1])
# add the bar chart to the new sheet

sheet.pictures.add(figure, name='bar chart', update=True)
# save the workbook

workbook.save()
workbook.close()
print("all files have been processed!")
app.quit()




