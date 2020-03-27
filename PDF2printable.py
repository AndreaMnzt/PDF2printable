import PyPDF4

from PIL import Image
import io
from docx import Document
import StringIO
from docx.shared import Cm, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
import sys
from insert_images import *

##Change this parameter to change the page format
#rows of images grid
n_rows = 3          
#columns of images grid
n_columns = 2
#width of the sidebar
side_bar = Cm(1.5)
#height of the top bar
top_bar = Cm(1.5)


document = Document()

#set shape to A4 paper
sections = document.sections
doc_height = sections[0].page_height = Cm(29.7) 
doc_width = sections[0].page_width  = Cm(21.0)
img_width = (doc_width - side_bar)/n_columns * 0.98
img_height = (doc_height - top_bar)/n_rows * 0.98


#remove margins of the page
for section in sections:
	section.top_margin = Cm(0)
	section.bottom_margin = Cm(0)
	section.left_margin = Cm(0)
	section.right_margin = Cm(0)
	#section.footer_distance = Cm(0)
	#section.header_distance = Cm(0)
	#section.gutter = Cm(0)
	#section.header.is_linked_to_previous = True
	#section.footer.is_linked_to_previous = True
	
sizes = [n_rows, n_columns, side_bar, top_bar, img_width, img_height]



if __name__ == "__main__":
	
	#errror if file not provided
	if(len(sys.argv) < 2):
		print("Usage: python PDF2printable.py <namefile.pdf>")
		exit(1)
	

	
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
				
				insert_images(document, sizes, imgs_per_page, right_page)
				
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
	# NOTE: just a repetition of the code above
	if len(imgs_per_page) != 0: 
			insert_images(document, sizes, imgs_per_page, right_page)	
				
			#since this is supposed to be the last page i don't need to flip it	
			#right_page = not right_page
				
			#clear memory		
			imgs_per_page = []


#save the docx file
document.save(sys.argv[1].replace(".pdf", ".docx"))
