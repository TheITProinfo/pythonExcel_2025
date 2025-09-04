import os
import xlwings as xw
print("current working directory: ",os.getcwd())
# get current file path
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# change working directory to current file path
os.chdir(current_path)
print("current working directory: ",os.getcwd())
app = xw.App(visible=True, add_book=False)
workbook = app.books.open('sum_total.xlsx')
# get all sheet names
sheets = workbook.sheets
print("sheets: ",sheets)
# loop through all sheets and rename the sheet name
for i in sheets:
    print("old sheet name: ",i.name)
    sheets[i].name = sheets[i].name .replace('销售', ' ')
workbook.save('sum_total_renamed.xlsx')
print("new sheet names: ",workbook.sheets)
workbook.close()
app.quit()
