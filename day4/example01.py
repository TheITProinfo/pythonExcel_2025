# utf-8
# -*- coding: utf-8 -*-
# this is example batch creating excel workbook
import os
import xlwings as xw
print("current working directory: ",os.getcwd())
# get current file path
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# change working directory to current file path
os.chdir(current_path)
print("current working directory: ",os.getcwd())

## bathch create excel workbook
app = xw.App(visible=True, add_book=False)
for i in range(1,10):
    workbook = app.books.add()
    workbook.save(f'example01_{i}.xlsx')
    workbook.close()
app.quit()





