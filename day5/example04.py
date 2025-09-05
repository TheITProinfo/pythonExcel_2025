# utf8  
# author: <NAME>
# date: 2021-01-04
# description: example04 for day5 get filtered data from excel and create new excel file
# version: 1.0
# usage: python example04.py
import os
import xlwings as xw
import pandas as pd
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# change working directory to file path
os.chdir(current_path)  
print("current working directory: ",os.getcwd())
# call excel
app = xw.App(visible=True, add_book=False)
# open excel file
workbook = app.books.open('purchase.xlsx')
# get all sheets in the workbook
sheets = workbook.sheets
data=[]
for sheet in sheets:
    print("sheet name: ",sheet.name)
    # get the value in the used range
    # value=sheet.range('A1').expand('table').value
    # print("data in used range: ",value)
    values=sheet.range('A1').expand('table').options(pd.DataFrame, index=False).value
    print("data in used range: ",values)
    filtered_data=values[values['采购物品']=='复印纸']
    if not filtered_data.empty:
        data.append(filtered_data)
# create new excel file
new_workbook = app.books.add()
# add new sheet to the workbook
new_sheet = new_workbook.sheets.add('printer paper', after=new_workbook.sheets[-1])
# write data to new sheet
new_sheet.range('A1').value = pd.concat(data, ignore_index=False)
# save new excel file
new_workbook.save('purchase_new.xlsx')
new_workbook.close()
print("all files have been processed!")
app.quit()
