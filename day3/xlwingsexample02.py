import xlwings as xw
# open an existing workbook
app = xw.App(visible=True, add_book=False)
workbook = app.books.open('example01.xlsx')
# get the first sheet
worksheet = workbook.sheets[0]
worksheet.range('A1').value = 'Hello, xlwings!'
# add new sheet
new_sheet = workbook.sheets.add('NewSheet')
# save the workbook
workbook.save()
# close the workbook and app
workbook.close()
app.quit()
