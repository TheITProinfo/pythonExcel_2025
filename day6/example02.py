# utf8
# author: <NAME>
# date: 2021-01-04
# description: Example for day 6, part 1, filtering excel data by column
import os
import xlwings as xw
import pandas as pd
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# call excel
app = xw.App(visible=True, add_book=False)
wb = app.books.open(current_path + "/purchase.xlsx")
sheets = wb.sheets
table_list=[]
for i, j in enumerate(sheets):
    print("sheet name: ",j.name)
    # get the value in the used range and convert to dataframe
    values= j.range("A1").expand('table').options(pd.DataFrame).value
    print(values)
    
    data=values.reindex(columns=['采购物品','采购日期','采购数量','采购金额'])
    print(data)
    #table=table.append(data,ingore_index=True)
    table_list.append(data)
    table=pd.concat(table_list,ignore_index=True)
    print(table)
    # group data by 采购物品 and sum 采购数量 and 采购金额
table=table.groupby(by='采购物品',as_index=False)
print(table)
# save data to new workbook
new_workbook = app.books.add()
for idx,group in table:
    print(idx)
    print(group)
    new_sheet = new_workbook.sheets.add(idx)
    new_sheet['A1'].options(index=False).value = group
    last_cell=new_sheet['A1'].expand('table').last_cell
    last_row=last_cell.row
    last_col=last_cell.column
    last_column_letter=chr(64+last_col)
    sum_cell_name='{}{}'.format(last_column_letter,last_row+1)
    sum_last_row_name='{}{}'.format(last_column_letter,last_row)
    formula='=SUM({}2:{}{})'.format(last_column_letter,last_column_letter,sum_last_row_name)
    new_sheet[sum_cell_name].value=formula
    new_sheet.autofit()
new_workbook.save(current_path + "/purchase_report.xlsx")
new_workbook.close()
print("all files have been processed!")
app.quit()