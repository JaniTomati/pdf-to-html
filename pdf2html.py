#!/usr/bin/env python3

import os
import winerror
from win32com.client.dynamic import Dispatch

from win32com.client.dynamic import ERRORS_BAD_CONTEXT


def convert_pdf_to_html(input, output):
    """ Convert a pdf to html by invoking Adobe Acrobat DC """
    src = os.path.abspath(input) # absolute path  

    app = Dispatch("AcroExch.AVDoc")
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

    print(os.path.join("converted", output_file))

    ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL) # avoiding "Not implemented" error 
    convert_pdf_to_html(input_file, output_file)

    # win32api.Beep(500, 3000)

    # speaker = win32com.client.Dispatch("SAPI.SpVoice")
    # speaker.Speak("It works. Hoorah!")


if __name__ == "__main__":
    main()