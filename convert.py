import os
from win32com import client
from shutil import copy

"""
This is the convert script:
	doc -> docx
"""

# Put all the .doc/.docx file need to be process in this path
path = r'H:/data/todo/all_doc'
# Specify the output path
output_path = r'H:/data/todo/all_doc/converted'


if not os.path.exists(output_path):
	os.mkdir(output_path)

def doc_to_docx(path):
	save_path = output_path + '/' + os.path.splitext(os.path.split(path)[1])[0] + ".docx"
	#print("Save to "+save_path)

	print("[*] Converting file from %s to %s" % (path,save_path))
	try:
		word = client.Dispatch('Word.Application')
		doc = word.Documents.Open(path)
		doc.SaveAs(save_path,16)
		doc.Close()
		word.Quit()
	except Exception as e:
		print("[-] Convert file %s failed" % path)
		print(e)


def main():
	for filename in os.listdir(path):
		filename = os.path.join(path,filename)
		# Convert all the .doc file to .docx
		if filename.split('.')[-1] == "doc":
			#print(filename)
			doc_to_docx(filename)
		# Pass all the .docx file to the output path
		elif filename.split('.')[-1] == "docx":
			try:
				copy(filename,output_path)
			except Exception as e:
				print(e)


if __name__ == '__main__':
	main()