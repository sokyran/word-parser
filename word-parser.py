import win32com.client
import os
import sys
from pathlib import Path

word = win32com.client.Dispatch("Word.Application")
word.visible = False

def collect_data(folder, write_file):
	for root, dirs, files in os.walk(folder):
		for i in files:
			if str(i).endswith('.doc') | str(i).endswith('.docx'):
				wb = word.Documents.Open(os.path.join(root, i))
				doc = word.ActiveDocument
				write_file.write('\n\n' + str(i).upper() + '\n\n')
				write_file.write(doc.Range().Text)
				print('\nDone')
				doc.Close()

def main():
	#'C:\\Users\\yurii\\Desktop\\Me\\Course 2\\Пархом\\Лабораторні роботи ОС\\admin_redak\\'
	folder = input("Enter folder location:\n") + "\\"
	file_name = folder.split('\\')[-2]
	with open(folder + f"\\{file_name}_collected.txt", 'w+', encoding='utf-8') as f:
		collect_data(folder, f)
	word.Quit()

main()