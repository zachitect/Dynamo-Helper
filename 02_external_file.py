import os
from os import listdir
from os.path import isfile, join

def txt_reader(file_path):
    file = open(file_path, "r")
    content = file.read()
    file.close()
    return content

def file_finder(dir_path):
	vptxt_dict = []
	for path in listdir(dir_path):
	    file_name = None
	    full_path = join(dir_path, path)
	    if isfile(full_path):
	        file_split = os.path.splitext(path)
	        if len(file_split) == 2:
	            if file_split[1].lower() == ".vptxt":
	                file_name = file_split[0]
	    if file_name != None:
	        vptxt_content = txt_reader(full_path)
	        vptxt_dict.append([file_name, vptxt_content])
	return vptxt_dict

# ----- Dynamo Input -----
if IN[0] == False:
    sys.exit("Operation Aborted")
folder_path = IN[1]
vptxt_entries = file_finder(folder_path)

# ----- Dynamo Output -----
OUT = vptxt_entries
