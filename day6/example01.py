# utf8
# author: <NAME>
# date: 2021-01-04
# description: Example for day 6, part 1, sorting excel data by column
import os
import xlwings as xw
import pandas as pd
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# call excel
app = xw.App(visible=True, add_book=False)
wb = app.books.open(current_path + "/product_sales_total.xlsx")
sheets = wb.sheets



for sheet in sheets:



    print("sheet name: ",sheet.name)
    # get the value in the used range and convert to dataframe
    #data = sheet.range("A1").options(pd.DataFrame, expand='table').value
    data= sheet.range("A1").expand('table').options(pd.DataFrame).value
    # print(data_list)
    print(data)
    #data.sort_values(by='销售利润')
    data = data.sort_values(by='销售利润', ascending=True) 
    print(data)
    # write data to new sheet
    sheet['A1'].options(index=False).value = data
    # sheet.range('A1').value = data
   
    sheet.autofit()
wb.save()
wb.close()
print("all files have been processed!")
app.quit()



