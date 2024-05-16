import pandas
import os
from tkinter.filedialog import askdirectory

def get_files(dir=None,return_dir = False):
    list = []
    for file in os.listdir(dir):
        if file.endswith(".pdf"):
            print(file)
            file_path = f"{file}"
            if return_dir == True:
                list.append(str(dir+"/"+file_path))
            else:
                list.append(str(file_path))
    return (list)
get_files(dir= askdirectory, return_dir= True)



import pandas
import os
from tkinter.filedialog import askdirectory

def get_files(return_dir=False):
    dir = askdirectory()  # Call askdirectory inside the function
    if not dir:
        return []  # Handle case where user cancels the dialog

    list_of_files = []
    for file in os.listdir(dir):
        if file.endswith(".pdf"):
            print(file)
            file_path = f"{dir}/{file}"  # Construct full path using the directory
            if return_dir:
                list_of_files.append(str(file_path))
            else:
                list_of_files.append(str(file))  # Just the filename

    return list_of_files

# Example usage
pdf_files = get_files(return_dir=True)  # Get full file paths
for file in pdf_files:
    print(file)


 
         