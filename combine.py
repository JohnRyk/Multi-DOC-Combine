# -*- coding: utf-8 -*-
"""
This is the combine script
Preinstall require: WPS or OFFICE
This script using the office api to combine the .docx file.
"""

import os
from os.path import abspath
from win32com import client

def main(files, final_docx):
	# Gain an office instance
	word = client.gencache.EnsureDispatch("Word.Application")
	word.Visible = True
	# New a blank document file
	new_document = word.Documents.Add()
	for fn in files:
		# Open all the file under process copy the content to Clipboard then close it.
		fn = abspath(fn)
		temp_document = word.Documents.Open(fn)
		word.Selection.WholeStory()
		word.Selection.Copy()
		temp_document.Close()
		# Append to the file
		new_document.Range()
		word.Selection.Delete()
		word.Selection.Paste()
	# Save the final document file and close the office instance
	new_document.SaveAs(final_docx)
	new_document.Close()
	word.Quit()
	
def find_file(path, ext, file_list=[]):
	dir = os.listdir(path)
	for i in dir:
		i = os.path.join(path, i)
		if os.path.isdir(i):
			find_file(i, ext, file_list)
		else:
			if ext == os.path.splitext(i)[1]:
				file_list.append(i)
	return file_list

if __name__ == '__main__':
	# cd /d H:/data/todo/all_doc/converted (If all you document are .docx you don't need the convert step, Just put this script into you folder than run it )

	# The output directory
	if not os.path.exists("./output"):
		os.mkdir("./output")
	dir_path = r"./" # Current directory 
	ext = ".docx"
	file_list = find_file(dir_path, ext)
	main(file_list,r"./output/result.docx")
	print(file_list)
