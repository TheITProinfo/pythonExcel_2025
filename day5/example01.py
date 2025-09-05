# utf-8
# -*- coding: utf-8 -*-
# this is example batch creating excel workbook
import os
import xlwings as xw
# get current path
cur_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",cur_path)
# change working directory to current file path
os.chdir(cur_path)
print("current working directory: ",os.getcwd())
# call excel
app = xw.App(visible=True, add_book=False)
# open existing workbook
wb = xw.Book('info.xlsx')
for sheet in wb.sheets:
    print("sheet name: ",sheet.name)
    # get all data in the used range
    value=sheet.range('A1').expand('table')
    # adjiust column width
    value.cloumn_width = 15
    # adjust row height
    value.row_height = 18
wb.save()
wb.close()
app.quit()