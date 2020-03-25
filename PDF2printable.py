import PyPDF4

from PIL import Image
import io
from docx import Document
import StringIO
from docx.shared import Cm, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
import sys

#change this parameter to change the page format
n_rows = 3
n_columns = 2
side_bar = Cm(2)

document = Document()

if __name__ == "__main__":
	
	
	
	
	#get tables in docx document
	tables = document.tables
	
	#set shape to A4 paper
	sections = document.sections
	img_height = sections[0].page_height = Cm(29.5) 
	img_width = sections[0].page_width  = Cm(20.8)

	#errror if file not provided
	if(len(sys.argv) < 2):
		print("Usage: python PDF2printable.py <namefile.pdf>")
		exit(1)
	
	#remove margins of the page
	for section in sections:
		section.top_margin = Cm(0)
		section.bottom_margin = Cm(0)
		section.left_margin = Cm(0)
		section.right_margin = Cm(0)
	

	
	input1 = PyPDF4.PdfFileReader(open(sys.argv[1], "rb"))
	
	#n of pages in pdf document
	n_pages = input1.getNumPages()
	
	#n of images per docx page
	n_imgs_per_doc_page = n_rows*n_columns
	
	#list for images data
	imgs_per_page = []
	
	#flag to check if page is right or left
	right_page = True
	
	#for every page in the pdf get the images and if n_imgs_per_doc_page are in the list add them to a docx page 
	for page_number in range(n_pages):
		
		#if the number of slide per page is reached insert them in a docx page
		if len(imgs_per_page) == n_imgs_per_doc_page:
				
				#start by a 0x(n_columns+1) grid
				table = document.add_table(rows=0, cols=n_columns+1)


				if right_page: #put the border in the left
					table.columns[0].width = side_bar
					for i in range(1,len(table.columns)):
						table.columns[i].width = (img_width-side_bar)/n_columns
					
				else: #put the border in the right
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
		
				#index for the column in the grid to insert the image
				col = 0
				
				#add a row to the table
				row_cells = table.add_row().cells
				
				#for every image
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
				
				
				#"flip" page for the next page
				right_page = not right_page
				
				#create a blank new page
				document.add_page_break()
					
				#remove the images in the images buffer
				imgs_per_page = []
				
				
		#for every image in the pdf file send it to the images butter (imgs_per_page)
		#credit to stackoverflow
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
					img = Image.frombytes(mode, size, data)
					imgs_per_page.append(img)
					
					#to save the image
					#img.save(obj[1:] + ".png")
					
				elif xObject[obj]['/Filter'] == '/DCTDecode':
					
					imgs_per_page.append(data)
					
					#to save the image
					#img = open(obj[1:] + ".jpg", "wb")
					#img.write(data)
					#img.close()
				elif xObject[obj]['/Filter'] == '/JPXDecode':
					
					imgs_per_page.append(data)
					
					#to save the image
					#img = open(obj[1:] + ".jp2", "wb")
					#img.write(data)
					#img.close()
					
	#if there are still images in the butter put them in a new page by repeating the above scrpt
	# NOTE: just a repetition of the code above, will be converted in a function in the future (maybe)
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


#save the docx file
document.save(sys.argv[1].replace(".pdf", ".docx"))
