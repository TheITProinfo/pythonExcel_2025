import os
# get the current working directory
print("current working directory is:",os.getcwd())
# create a new directory
# os.mkdir("test_folder")
# change the current working directory
os.chdir("test_folder")
# get the current working directory
print("current working directory is:",os.getcwd())
# rename the files
old_name="newfile.txt"
new_name="renamedfile.txt"
os.rename(old_name,new_name)

