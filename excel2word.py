from docx import Document
from docx.shared import Inches
import os, glob, tkinter as tk
from tkinter import filedialog
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import io

# file  = openpyxl.load_workbook('exceltest.xlsx',data_only = True)

# sheet = file['Roller Coater Cleaning']

# image_loader = SheetImageLoader(sheet)

# image = image_loader.get('B14')

# image.show()

# imageByte = io.BytesIO()

# image.save(imageByte, format='BMP')

# data = imageByte.getvalue()


# doc = Document("Testing.docx")

# doc.tables[1].add_row().cells  

# doc.save("Testing.docx")


# cell = doc.tables[1].rows[1].cells[1].paragraphs[0].add_run()

# cell.add_picture(imageByte, width = Inches(2), height = Inches(2))

# doc.save('Testing.docx')

# counter = 0

# for row in sheet.iter_rows(1,20, values_only = True):
# 	counter += 1
# 	print(row[0])
# 	# if row[0] == 1:
# 	# 	break

# print(counter)


	

"""
Steps to completion

1) Make sure that text can be transfered to the cells in word table in

2) check function for adding rows

3) start putting it all together. 



"""

file  = openpyxl.load_workbook(filedialog.askopenfilename(title = "Select the excel file to be transferred from."), data_only=True)

sheet = file.sheetnames

sheet = file[sheet[0]]

# print(sheet.max_row)


image_loader = SheetImageLoader(sheet)

worder = filedialog.askopenfilename(title = "Select the word file to be transferred to.")

doc = Document(worder)

table = doc.tables[1]

counter = 0

for row in sheet.iter_rows(1, 20, values_only= True):

	counter += 1


	if row[0] == 1:
		break


for row in sheet.iter_rows(counter, sheet.max_row,  values_only= True):

	rowcells = table.add_row().cells

	if type(row[0]).__name__ == 'str':
		
		rowcells[0].paragraphs[0].add_run(row[0])

	else :

		rowcells[2].paragraphs[0].add_run(row[2])

		rowcells[3].paragraphs[0].add_run('None')

		rowcells[4].paragraphs[0].add_run(row[3])

		try:

			image = image_loader.get("B" + str(counter))

			imageByte = io.BytesIO()

			image.save(imageByte, format= "BMP")

			run = rowcells[1].paragraphs[0].add_run()

			run.add_picture(imageByte, width = Inches(2), height = Inches(2))

		except :
			next


		

	counter +=1


		


doc.save(worder)

	









