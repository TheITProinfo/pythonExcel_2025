import os
# get current file path
current_path = os.path.dirname(os.path.abspath(__file__))
print("current file path: ",current_path)
# change working directory to current file path
os.chdir(current_path)
print("current working directory: ",os.getcwd())

# list all files in current directory
files = os.listdir(current_path)
print("files in current directory: ",files)
for file in files:
    print(file)