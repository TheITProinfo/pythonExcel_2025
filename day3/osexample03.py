import os
file_anme="renamedfile.txt"
# split the file name and extension
file_name, file_extension = os.path.splitext(file_anme)
print("file name:",file_name)
print("file extension:",file_extension)

