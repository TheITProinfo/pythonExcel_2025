# utf8
# author: <NAME>
# date: 2021-01-04
# this script is to demonstrate how to read multiple excel files in a folder, process the data, and write the results back to a new sheet in each file
import os
import xlwings as xw
import pandas as pd
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# call excel
app = xw.App(visible=True, add_book=False)
# open excel file


file_path = os.path.join(current_path, "purchase.xlsx")
# open file alwasys use absolute path
workbook = app.books.open(file_path )



# get all sheets in the workbook



sheets = workbook.sheets



data=[]



for sheet in sheets:



    print("sheet name: ",sheet.name)



    # get the value in the used range



    value_area=sheet.range('A1').expand('table')
    data=value_area.options(pd.DataFrame, index=False, header=True).value
    print(data)
    sums=data['采购金额'].sum()
    print("采购金额总和: ",sums)
    # get the column number of 采购金额
    col_num=data.columns.get_loc('采购金额')+1
    # get last row number
    last_row=value_area.last_cell.row
    print("last row: ",last_row)
    sheet.range((last_row+1,col_num)).value=sums

workbook.save()
workbook.close()
print("all files have been processed!")
app.quit()

    



    