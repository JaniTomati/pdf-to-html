#!/usr/bin/env python3


""" 
    python3 pdf2html.py -d [input-pdf-dir] -o [output-html-dir]

    Invokes the Adobe Acrobat DC to transform PDFs into HTML files.
"""


import os
import sys
import argparse
import winerror
from win32com.client.dynamic import Dispatch

from win32com.client.dynamic import ERRORS_BAD_CONTEXT


def set_up_parser():
    """ Set up arguments for command line argument parsing """
    parser = argparse.ArgumentParser()
    parser.add_argument("-d", "--pdf-dir", required=True, help="add path to input directory containing pdfs.")
    parser.add_argument("-o", "--html-dir", required=True, help="add path to the output directory containing html files.")
    args = parser.parse_args()
    return args


def parse_arguments(arguments):
    """ Parse the arguments that have been provided by the user via command line """
    input_path = arguments.pdf_dir
    output_path = arguments.html_dir

    if not os.path.isdir(input_path):
        print("Error: Input directory does not exist.\nStopping execution.")
        sys.exit(1)
    
    if not os.path.isdir(output_path):
        print("Error: Output directory does not exist.\nStopping execution.")
        sys.exit(1)

    return input_path, output_path


def convert_pdf_to_html(input, output):
    """ Convert a pdf to html by invoking Adobe Acrobat DC """
    src = os.path.abspath(input) 

    app = Dispatch("AcroExch.AVDoc") # Adobe Acrobat
    app.Open(src, src)
    pdDoc = app.GetPDDoc()
    jsObject = pdDoc.GetJSObject()
    jsObject.SaveAs(os.path.join("converted", output), "com.adobe.acrobat.html")

    pdDoc.Close()
    app.Close(True)
    del pdDoc


def main():
    input_file = "vizsnippets_munzner2021.pdf"
    output_file  = "vizsnippets_munzner2021.html"

    # parse arguments
    arguments = set_up_parser()
    input_path, output_path = parse_arguments(arguments)

    # convert pdf2html
    ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL) # avoiding "Not implemented" error 
    #convert_pdf_to_html(input_file, output_file)


if __name__ == "__main__":
    main()