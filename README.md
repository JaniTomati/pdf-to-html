# PDF-to-HTML converter
Python script that invokes Adobe Acrobat to automatically convert PDF files to HTML files using Windows. 

## Dependencies 
- Adobe Acrobat DC 
- winerror 
- [pywin32](https://pypi.org/project/pywin32/) 

# Usage 
Run the script using ``python3 pdf2html.py`` and provide the following two arguments 
  * ``-d, --pdf-dir``: directory that holds a number of pdf files that are to be converted to html. 
  * ``-o, --html-dir``: directory where the html conversions are to be stored. 

Adjust the code as necessary depending on your file system and structure. 
