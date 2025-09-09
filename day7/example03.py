# utf8  
# author: <NAME>
# date: 2021-01-04
# this is an example to read multiple excel files in a folder, process the data, and write the results back to a new sheet in each file
import os
import xlwings as xw
import pandas as pd
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# join path
file_path = os.path.join(current_path, "sales_total_by_product")
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
            max=data['销售利润'].max()
            print("max: ",max)
            min=data['销售利润'].min()
            print("min: ",min)
            avg=data['销售利润'].mean().round(2)
            print("avg: ",avg)
            # get the last row number
            last_row=sheet.range("A1").expand('table').last_cell.row
            print("last row: ",last_row)
            # write data to new sheet
            
            sheet['A'+str(last_row+4)].value='maximum'
            sheet['B'+str(last_row+4)].value='minimum'
            sheet['C'+str(last_row+4)].value='average'
            sheet['A'+str(last_row+5)].value=max
            sheet['B'+str(last_row+5)].value=min
            sheet['C'+str(last_row+5)].value=avg
            sheet.autofit()
        workbook.save()
        workbook.close()
print("all files have been processed!")
            
app.quit()
