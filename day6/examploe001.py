# utf8
# author: <NAME>
# date: 2021-01-04
# description: Example for day 6, part 1, sorting excel data by column

import os
import xlwings as xw
import pandas as pd

# Current script path
current_path = os.path.dirname(os.path.abspath(__file__))
print("Current file path:", current_path)

# Open Excel
app = xw.App(visible=True, add_book=False)
wb = app.books.open(os.path.join(current_path, "product_sales_total.xlsx"))

for sheet in wb.sheets:
    print("Processing sheet:", sheet.name)

    # Read sheet into DataFrame
    data = sheet.range("A1").expand("table").options(pd.DataFrame).value
    print("Original data:\n", data)

    # Sort ascending (small → large) by 销售利润
    data = data.sort_values(by="销售利润", ascending=True)
    print("Sorted data:\n", data)

    # Write back to Excel
    sheet["A1"].options(index=False).value = data
    sheet.autofit()

# Save & close
wb.save()
wb.close()
print("All files have been processed!")
app.quit()
