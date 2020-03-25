import PyPDF4

from PIL import Image
import io
from docx import Document
import StringIO
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
import sys

#change this parameter to change the page format
n_rows = 3
n_columns = 2
side_bar = Cm(2)

document = Document()

if True:
	
	
	
	

	tables = document.tables

	sections = document.sections
	img_height = sections[0].page_height = Cm(29.5) 
	img_width = sections[0].page_width  = Cm(20.8)

	
	if(len(sys.argv) < 2):
		print("Usage: python PDF2printable.py <namefile.pdf>")
		exit(1)
	for section in sections:
		section.top_margin = Cm(0)
		section.bottom_margin = Cm(0)
		section.left_margin = Cm(0)
		section.right_margin = Cm(0)
	

	
	input1 = PyPDF4.PdfFileReader(open(sys.argv[1], "rb"))
	n_pages = input1.getNumPages()
		
	n_imgs_per_doc_page = n_rows*n_columns
	imgs_per_page = []
	right_page = True
	for page_number in range(n_pages):
		
		#if the number of slide per page is reached insert them in a docx page
		if len(imgs_per_page) == n_imgs_per_doc_page:
				
				table = document.add_table(rows=0, cols=n_columns+1)
				if right_page:
					table.columns[0].width = side_bar
					#table.columns[1].width = table.columns[2].width = (img_width-side_bar)/n_columns
					for i in range(1,len(table.columns)):
						table.columns[i].width = (img_width-side_bar)/n_columns
					
				else:
					table.columns[-1].width = side_bar
					for i in range(len(table.columns)-1):
						table.columns[i].width = (img_width-side_bar)/n_columns
					
				table.style = "Table Grid"
				table.autofit = False
				table.allow_autofit = False

				#if right_page:
				#	table.alignment = WD_TABLE_ALIGNMENT.LEFT
				#else: 
				#	table.alignment = WD_TABLE_ALIGNMENT.LEFT
		
				
				col = 0
				row_cells = table.add_row().cells
				for i in range(0,len(imgs_per_page)):
					
					#get the image to print in the table this cycle\
					image = imgs_per_page[i]
										
					#every n_column image change row by creating a new row 
					if(col == n_columns):
						col = 0
						row_cells = table.add_row().cells
							
					#change border accoring to right or left page
					if right_page:
						actual_col = col+1
					else:
						actual_col = col


					#create paragraph within the table 
					paragraph = row_cells[actual_col].paragraphs[0]
					cell_paragraph_format = paragraph.paragraph_format
					cell_paragraph_format.left_indent = Cm(0)
					cell_paragraph_format.right_indent = Cm(0)
					
					#set alignment of picture inside the table
					#if right_page:
					paragraph.alignment = WD_TABLE_ALIGNMENT.LEFT
					#else: 
					#	paragraph.alignment = WD_TABLE_ALIGNMENT.LEFT
					
					#add image to table
					run = paragraph.add_run()
					run.add_picture(StringIO.StringIO(image), width = (img_width-side_bar)/n_columns*0.98 , height = img_height/n_rows*98/100)
					
					col += 1
				
				
				
				right_page = not right_page
				
				document.add_page_break()
					
				imgs_per_page = []
				
		page0 = input1.getPage(page_number)
		xObject = page0['/Resources']['/XObject'].getObject()

		for obj in xObject:
			if xObject[obj]['/Subtype'] == '/Image':
				size = (xObject[obj]['/Width'], xObject[obj]['/Height'])
				data = xObject[obj].getData()
				if xObject[obj]['/ColorSpace'] == '/DeviceRGB':
					mode = "RGB"
				else:
					mode = "P"

				if xObject[obj]['/Filter'] == '/FlateDecode':
					print(1)
					img = Image.frombytes(mode, size, data)
					#img.save(obj[1:] + ".png")
					imgs_per_page.append(img)
				elif xObject[obj]['/Filter'] == '/DCTDecode':
					#img = open(obj[1:] + ".jpg", "wb")
					
					imgs_per_page.append(data)
					
					#img.write(data)
					#img.close()
				elif xObject[obj]['/Filter'] == '/JPXDecode':
					print(3)
					#img = open(obj[1:] + ".jp2", "wb")
					#img.write(data)
					#img.close()
					
					imgs_per_page.append(data)
					
					
	if len(imgs_per_page) != 0:
			
			table = document.add_table(rows=0, cols=n_columns+1)
			if right_page:
				table.columns[0].width = side_bar
			#table.columns[1].width = table.columns[2].width = (img_width-side_bar)/n_columns
				for i in range(1,len(table.columns)):
					table.columns[i].width = (img_width-side_bar)/n_columns
					
			else:
				table.columns[-1].width = side_bar
				for i in range(len(table.columns)-1):
					table.columns[i].width = (img_width-side_bar)/n_columns
					
			table.style = "Table Grid"
			table.autofit = False
			table.allow_autofit = False

				#if right_page:
				#	table.alignment = WD_TABLE_ALIGNMENT.LEFT
				#else: 
				#	table.alignment = WD_TABLE_ALIGNMENT.LEFT
		
				
			col = 0
			row_cells = table.add_row().cells
			for i in range(0,len(imgs_per_page)):
					
				#get the image to print in the table this cycle\
				image = imgs_per_page[i]
										
				#every n_column image change row by creating a new row 
				if(col == n_columns):
					col = 0
					row_cells = table.add_row().cells
						
				#change border accoring to right or left page
				if right_page:
					actual_col = col+1
				else:
					actual_col = col


				#create paragraph within the table 
				paragraph = row_cells[actual_col].paragraphs[0]
				cell_paragraph_format = paragraph.paragraph_format
				cell_paragraph_format.left_indent = Cm(0)
				cell_paragraph_format.right_indent = Cm(0)
					
				#set alignment of picture inside the table
				#if right_page:
				paragraph.alignment = WD_TABLE_ALIGNMENT.LEFT
				#else: 
				#	paragraph.alignment = WD_TABLE_ALIGNMENT.LEFT
			
				#add image to table
				run = paragraph.add_run()
				run.add_picture(StringIO.StringIO(image), width = (img_width-side_bar)/n_columns*0.98 , height = img_height/n_rows*98/100)
			
				col += 1
				
				
				
			right_page = not right_page
				
					
			imgs_per_page = []



document.save(sys.argv[1].replace(".pdf", ".docx"))
