import win32com.client
import os
import sys
from pathlib import Path

word = win32com.client.Dispatch("Word.Application")
word.visible = False

# folder = input('Enter path to folder with doc files: ')
# file_name = folder.split('\\')[-2]

# Previous folder settings
# folder + '\\{}.txt'.format(file_name)


def handle_folder(folder, f):
	entries = Path(folder)
	for i in entries.iterdir():
		if i.is_dir():
			handle_folder(i, f)
			print(f)
			print(i, ' is folder')
		elif str(i).endswith('.doc') | str(i).endswith('.docx'):
			handle_doc(i, f)


def handle_doc(docum, file):
	print(docum)
	wb = word.Documents.Open(str(docum))
	doc = word.ActiveDocument
	file.write('\n\n' + str(docum).upper() + '\n\n')
	file.write(doc.Range().Text)
	print('\nDone')
	doc.Close()


folder = 'C:\\Users\\yurii\\Desktop\\Me\\Course 2\\Базы данных\\Материалы\\'
with open(folder + '\\all_lections.txt', 'wb+', encoding='utf-8') as f:
	handle_folder(folder, f)

word.Quit()
