# utf8
# author: <NAME>
# date: 2021-01-04
# description: example day 7 creating pivot table
import os
import xlwings as xw
import pandas as pd
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
file_path = os.path.join(current_path, "sales_by_product")
# list all files in the foldercl
file_list = os.listdir(file_path)
print("files in folder: ", file_list)
# call excel
app = xw.App(visible=True, add_book=False)
# loop all files
for file in file_list:
    if file.startswith('~$'):
        continue
    if file.endswith('.xlsx'):
        workbook = app.books.open(os.path.join(file_path, file))
        # get all sheets in the workbook
        sheets = workbook.sheets
        for sheet in sheets:
            print("sheet name: ",sheet.name)
            # get the value in the used range and convert to dataframe
            data=sheet.range("A1").expand('table').options(pd.DataFrame, index=False, header=True).value
            print(data)
            pivottable=pd.pivot_table(data,index=['销售地区'],values=['销售金额'],columns='销售分部',aggfunc='sum')
            print(pivottable)
            # write data to current sheet
            sheet.range('k1').value=pivottable
            sheet.autofit()

        workbook.save()
        workbook.close()
print("all files have been processed!")
            
app.quit()            
    