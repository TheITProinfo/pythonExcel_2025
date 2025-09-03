import xlwings as xw
import os
# get current working directory
cwd = os.getcwd()
print(f'Current working directory: {cwd}')

# create a new workbook
app = xw.App(visible=True, add_book=True)
workbook = app.books.add()
workbook.save('example01.xlsx')
# close the workbook and app
workbook.close()
app.quit()
