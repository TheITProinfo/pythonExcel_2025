# utf8
# author: <NAME>
# date: 2021-01-04
# description: example02 for day5 adjust cell format
# version: 1.0
# usage: python example02.py
import os
import xlwings as xw
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)

file_path=os.path.join(current_path,'purchase')
print("file path: ",file_path)
# change working directory to file path
os.chdir(file_path)
print("current working directory: ",os.getcwd())
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
        sheets = workbook.sheets
        for sheet in sheets:
            print("sheet name: ",sheet.name)
            # get the last row number in the used range
            row_num = sheet.range('A1').current_region.rows.count
            print("last row number: ",row_num)
            sheet['A2:A{}'.format(row_num)].number_format = 'm/d'
            sheet['D2:D{}'.format(row_num)].number_format = '#,##0.0'
        workbook.save()
        workbook.close()
print("all files have been processed!")
app.quit()  