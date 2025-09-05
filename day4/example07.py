# utf-8
# -*- coding: utf-8 -*-
# this is example batch creating excel workbook
import os
import xlwings as xw

current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# change working directory to current file path
os.chdir(current_path)
print("current working directory: ",os.getcwd())

file_path = os.path.join(current_path, 'sales')
print("file path: ",file_path)
file_list = os.listdir(file_path)
# run app
app = xw.App(visible=True, add_book=False)
# open workbook
workbook = app.books.open('info.xlsx')

worksheets = workbook.sheets   # get all sheets in the workbook -info.xlsx
for i in file_list:  # target workbook
    if i.startswith('~$'):
        continue
    if i.endswith('.xlsx'):
        name = i.split('_')[0]
        print(name)
        new_file_path = os.path.join(file_path, i)
        # \\ open file with absolute path
        # \\d:\code\pythoncode\pythonExcel_2025\day4\sales\单肩包.xlsx
        workbooks = app.books.open(new_file_path)
        # workbooks for target workbook
        for j in worksheets:  # loop through all sheets in info.xlsx
           contents = j.range('A1').expand('table').value  # get all data in the sheet
           name=j.name # get sheet name
           workbooks.sheets.add(name, after=workbooks.sheets[-1]) # add new sheet to the target workbook
           workbooks.sheets[name].range('A1').value = contents # write data to new sheet
        workbooks.save()
        

app.quit()