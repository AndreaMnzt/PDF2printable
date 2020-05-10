#!/usr/bin/python

import PyPDF4
from PIL import Image
import io
from docx import Document
from docx.shared import Cm, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
import sys
from insert_images import *
from wand.image import Image as wImage

##Change this parameter to change the page format
#rows of images grid
n_rows = 3
#columns of images grid
n_columns = 2
#width of the sidebar
side_bar = Cm(1.5)
#height of the top bar
top_bar = Cm(.5)

#angle to rotate
#angle  = 90 feature in dev
#put new_page to false if blank pages appear
new_page = False

document = Document()

#set shape to A4 paper
sections = document.sections
doc_height = sections[0].page_height = Cm(29.7)
doc_width = sections[0].page_width  = Cm(21.0)
img_width = (doc_width - side_bar)/n_columns * 0.98
img_height = (doc_height - top_bar)/n_rows * .93


#remove margins of the page
for section in sections:
	section.top_margin = Cm(0)
	section.bottom_margin = Cm(0)
	section.left_margin = Cm(0.5)
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

	resolution = 200

	#for every page in the pdf get the images and if n_imgs_per_doc_page are in the list add them to a docx page
	for page_number in range(n_pages):

		#if the number of slide per page is reached insert them in a docx page
		if len(imgs_per_page) == n_imgs_per_doc_page:

				insert_images(document, sizes, imgs_per_page, right_page)

				#"flip" page for the next page
				right_page = not right_page

				#create a blank new page
				if new_page:
					document.add_page_break()

				#remove the images in the images buffer
				imgs_per_page = []

		#for every image in the pdf file send it to the images butter (imgs_per_page)
		page0 = input1.getPage(page_number)
		dst_pdf = PyPDF4.PdfFileWriter()
		dst_pdf.addPage(page0)
		pdf_bytes = io.BytesIO()
		dst_pdf.write(pdf_bytes)
		pdf_bytes.seek(0)
		page_img = wImage(file = pdf_bytes, resolution = resolution)
		page_img = page_img.make_blob('JPG')
		imgs_per_page.append(page_img)

		sys.stdout.flush()
		sys.stdout.write("\r{0}.".format("Converting page "+ str(page_number+1) + " of "+ str(n_pages)))

	#if there are still images in the butter put them in a new page by repeating the above scrpt
	# NOTE: just a repetition of the code above
	if len(imgs_per_page) != 0:
			insert_images(document, sizes, imgs_per_page, right_page)

			#since this is supposed to be the last page i don't need to flip it
			#right_page = not right_page

			#clear memory
			imgs_per_page = []

#save the docx file

filename = sys.argv[1]
document.save(filename.replace(".pdf", ".docx"))
print('')
print('Done. :) \nDocument saved as \'' + filename.replace(".pdf",".docx") + '\'.')
