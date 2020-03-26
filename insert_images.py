from PIL import Image
import io
from docx import Document
import StringIO
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT



def insert_images(document, sizes, imgs_per_page, right_page):
	n_rows, n_columns, side_bar, top_bar, img_width, img_height = sizes
	
	#get tables in docx document
	tables = document.tables
	
	#start by a 0x(n_columns+1) grid
	table = document.add_table(rows=1, cols=n_columns+1)
	table.rows[0].height = top_bar

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
					
		#get the image to print in the table this cycle
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
		run.add_picture(StringIO.StringIO(image), width = (img_width-side_bar)/n_columns*0.98 , height = (img_height-top_bar)/n_rows*98/100)
					
		col += 1
				
		
