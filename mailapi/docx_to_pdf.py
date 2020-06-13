# -*- coding: utf-8 -*-
"""
Created on Tue May 19 12:35:25 2020

@author: Bhanu joshi
"""

import os
import comtypes.client
from docx2pdf import convert
# import win32com.client
import pywin32_system32
import pythoncom
import win32com.client as client

# pythoncom.CoInitialize()
#
# strComputer = "."
# objWMIService = client.Dispatch("WbemScripting.SWbemLocator")
# objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
# colItems = objSWbemServices.ExecQuery("Select * from Win32_Process")


wdFormatPDF = 17

def convertFile(file):
    pythoncom.CoInitialize()
    in_file = os.path.abspath(file)
    out_file = os.path.abspath(file.replace(".docx", ".pdf").replace(".doc", ".pdf"))
    # word = win32com.client.DispatchEx("Word.Application")
    try:
        word = client.Dispatch("Word.application")
        wordDoc = word.Documents.Open(in_file, False, False, False)
        # wordDoc.SaveAs2(out_file, FileFormat=wdFormatPDF)
        wordDoc.SaveAs(out_file, FileFormat=wdFormatPDF)
        wordDoc.Close()
        #word = comtypes.client.CreateObject('Word.Application')
        # doc = word.Documents.Open(in_file)
        # doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        # doc.Close()
        word.Quit()
        print("Hii bhanu")
    except Exception as e:
        print(e)
    return out_file