# PDF-to-HTML converter
Python script that invokes Adobe Acrobat to automatically convert PDF files to HTML files using Windows. 

## Dependencies 
- Adobe Acrobat DC 
- winerror 
- [pywin32](https://pypi.org/project/pywin32/) 

## Usage 
Run the script using ``python3 pdf2html.py`` and provide the following two arguments 
  * ``-d, --pdf-dir``: directory that holds a number of pdf files that are to be converted to html. 
  * ``-o, --html-dir``: directory where the html conversions are to be stored. 

Adjust the code as necessary depending on your file system and structure. 

Use the ``convert_pdf_to_html`` function and provide the following parameters to convert a file
 * ``input``: A path to a valid PDF document. 
 * ``output``: A path to an output folder where the converted HTML file is saved.
 * ``name``: The name under which the HTML output file will be saved (the extension ``.html``is added automatically). 
