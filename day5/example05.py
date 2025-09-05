# utf8  
# author: <NAME>
# date: 2021-01-04
# description: example04 for day5 split data from excel and replace value
# version: 1.0
# usage: python example04.py
import os
import xlwings as xw
import pandas as pd
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
## get file path
file_path=os.path.join(current_path,'product_speci')
print("file path: ",file_path)
# change working directory to file path
os.chdir(file_path) 
print("current working directory: ",os.getcwd())
# list all files in current directory
file_list = os.listdir(file_path)
print("files in current directory: ",file_list)

# call excel
app = xw.App(visible=True, add_book=False)
# open excel file
for file in file_list:
    if file.startswith('~$'):
        continue
    if file.endswith('.xlsx'):
        workbook = app.books.open(file)
        # get all sheets in the workbook
        sheets = workbook.sheets['规格表']
        print("sheet name: ",sheets.name)
        values=sheets.range('A1').expand('table').options(pd.DataFrame, index=False).value
        print("data in used range: ",values)
        # split data
        new_values=values['规格'].str.split('*',expand=True)
        print("split data: ",new_values)
        values['length']=new_values[0]
        values['width']=new_values[1]
        values['height']=new_values[2]
        print("new data: ",values)
        # drop original column
        values.drop(columns=['规格'],inplace=True)
        print("after drop original column: ",values)
        # write data to new sheet
        sheets['A1'].options(index=False).value=values
        sheets.autofit()
        workbook.save()
        workbook.close()
print("all files have been processed!")
app.quit()
        
        
