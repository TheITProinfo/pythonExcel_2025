import os
# get the name of the operating system
print("operating system name is:",os.name)
# get current working directory
print("current working directory is:",os.getcwd())
# get the list of files and directories in the current directory
print("list of files and directories in the current directory is:",os.listdir())
# get the current user
print("current user is:",os.getlogin())
# get the current user home directory
print("current user home directory is:",os.path.expanduser('~'))
# change the current working directory
os.chdir("C:\Windows")
print("current working directory is:",os.getcwd())
# create a new directory
os.mkdir("newdir")

