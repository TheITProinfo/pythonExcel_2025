import os
import xlwings as xw
# get current file path
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# change working directory to current file path
os.chdir(current_path)
print("current working directory: ",os.getcwd())
# join path
file_path = os.path.join(current_path, 'product_sales')
print("file path: ",file_path)
# change working directory to file path
os.chdir(file_path)
print("current working directory: ",os.getcwd())
# list all files in current directory
file_list = os.listdir(file_path)
print("files in current directory: ",file_list)
old_book_name='销售表'
new_book_name='product_sales'
# loop through all files and rename the file name   
for file in file_list:
    if file.startswith('~$'):
        continue
    new_file = file.replace(old_book_name, new_book_name)
    new_file_path = os.path.join(file_path, new_file)
    old_file_path = os.path.join(file_path, file)
    os.rename(old_file_path, new_file_path)
    print(f'rename file: {file} to {new_file}')