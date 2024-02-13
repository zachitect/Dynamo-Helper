import os
from os import listdir
from os.path import isfile, join

#read txt file
def txt_reader(file_path):
    file = open(file_path, "r")
    content = file.read()
    file.close()
    return content

#list all files of specific extension from directory
def files_from_directory(dir_path, dot_ext):
    file_paths = []
    for path in listdir(dir_path):
        full_path = join(dir_path, path)
        if isfile(full_path): #confirm that the path is a file not a directory
	    file_split = os.path.splitext(path)
	    if len(file_split) == 2: #root path + extension name
	        if file_split[1].lower() == dot_ext.lower(): #check extension name
		        file_paths.append(full_path)
    return file_paths
