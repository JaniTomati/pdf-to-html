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
        print("Output directory does not exists.\nCreating new dir at " + output_path + ".")
        os.mkdir(output_path)

    return input_path, output_path


def convert_pdf_to_html(input, output, name):
    """ Convert a pdf to html by invoking Adobe Acrobat DC """
    src = os.path.abspath(input) 
    dest = os.path.abspath(output + "/" + name)

    app = Dispatch("AcroExch.AVDoc") # Adobe Acrobat
    app.Open(src, src)
    pdDoc = app.GetPDDoc()
    jsObject = pdDoc.GetJSObject()
    jsObject.SaveAs(os.path.join(dest, name + ".html"), "com.adobe.acrobat.html")

    pdDoc.Close()
    app.Close(True)
    del pdDoc


def convert_publications(input_path, output_path):
    """ Specific use case: convert a number of publications from PDF to HTML """
    #input_path = "D:/IEEE Vis+InfoVis+Vast+VisWeek/VisWeek/IEEE VIS 2021/pdfs"
    #output_path = "html_data/2021/"

    conv_cnt = 0
    for workshop_dir in os.listdir(input_path):
        for contribution_dir in os.listdir(input_path + "/" + workshop_dir):
            for filename in os.listdir(input_path + "/" + workshop_dir + "/" + contribution_dir):
                if filename.endswith(".pdf"):
                    file = input_path + "/" + workshop_dir + "/" + contribution_dir + "/" + filename
                    name = os.path.splitext(file)[0].split("/")[-1] # get file name without file extension

                    if os.path.isdir(output_path + "/" + name):
                        print("File", name, "is already converted.")
                    else:
                        convert_pdf_to_html(file, output_path, name)
                        conv_cnt += 1
                
                    if conv_cnt != 0 and conv_cnt % 10 == 0:
                        print("Processed", conv_cnt, "articles.")

    return conv_cnt


def main():
    # parse arguments
    arguments = set_up_parser()
    input_path, output_path = parse_arguments(arguments)

    # convert pdf2html
    ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL) # avoiding "Not implemented" error

    conv_cnt = convert_publications(input_path, output_path)
    print("Processed", conv_cnt, "articles in total.")
            

if __name__ == "__main__":
    main()