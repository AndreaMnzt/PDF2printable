# PDF2printable
A useful script to print pptx file made available as pdf. 

The pdf pages are resized in the grid of a docx file to maximize the space for printing.

**The script is useful for students to print a pptx file made available as a pdf file.**

<img src = "images/script_description.png" width  = 1000></img>

By passing your pdf file to the script a .docx file with all the pages (the slides in the pptx) in a grid is provided.
The space is maximized to print the slides as large as possible.
A lateral space is at the left/right of the images to make the print perfect to be inserted in a notebook. 

## Usage:
- save the pdf file in the folder of PDF2printable.py
- move with the teminal to the same folder and run:

<code>pyhton PDF2printable.py filename.pdf </code>

The ouput of the script is _filename.docx_ file in the filename.pdf folder.

## Install
<em>Intall the libraries</em>

<code>pip install pillow PyPDF4 python-docx wand</code>

<em>Download the script</em>
<li><code>git clone https://github.com/AndreaMnzt/PDF2printable.git</code></li>

<em>If you are using a Windows OS you have to install imageMagick library, you can find it at</em>
<a>http://docs.wand-py.org/en/latest/guide/install.html#install-imagemagick-on-windows</a>
 
 
## Options
At the moment the only way to customize the script is edit the following lines of code:

<code>n_rows = 3</code>

<code>n_columns = 2</code>

<code>side_bar = Cm(1.5)</code>

<code>top_bar = Cm(1.5)</code>

- _n_rows_ specifies the rows of the images grid in the output docx file.
- _n_columns_ specifies the columns of the images grid in the output docx file.
- _side_bar_ specifies the width of the sidebar to make the page fit a notebook page (or to insert it in a 4-ring binder).
The value of _side_bar_ can be specified in Cm(_centimeter value_) and Inches(_inches value_). Set it to Cm(0) to remove the sidebar and make the image grid full page. 

Note 1: this is implemented as a blank column in the images table.
- _top_bar_ specifies the height of the header of the page. 

Note 1: this is implemented as a blank row in the images table.

Note 2: top_bar is currently not represented in the description image above.
### Note:
- Since the docx has a quasi fullsreen table of images, to print the document the printer should be able to print without margins.
- The script is just a (working) sketch.

### Credits
- Stackoverflow for some code snippets
