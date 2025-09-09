# utf8  
# author: <NAME>
# date: 2021-01-04
# description: example day 7 sales total group by sales area, then sum sales amount
import os
import xlwings as xw

import pandas as pd

current_path = os.path.dirname(os.path.abspath(__file__))

print("current file path: ",current_path)

file_path = os.path.join(current_path, "sales total")
# list all files in the foldercl
file_list = os.listdir(file_path)
print("files in folder: ", file_list)
# call excel
app = xw.App(visible=False, add_book=False)
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
            data = sheet.range("A1").expand('table').options(pd.DataFrame, index=False, header=True).value
            print(data)
            data['销售利润']=data['销售利润'].astype('float')
            # group by sale area and sum sales amount
            result=data.groupby('销售区域')['销售利润'].sum()
            print(result)
            # write data to new sheet
            # creat new sheet
            if 'Total' in [s.name for s in sheets]:
                sheet=sheets['total']
                sheet.clear()
            else:
                sheet=workbook.sheets.add('total')
            sheet['A1'].options(index=False).value=result
            sheet.autofit()
        workbook.save()
        workbook.close()
print("all files have been processed!")
app.quit()
            
            