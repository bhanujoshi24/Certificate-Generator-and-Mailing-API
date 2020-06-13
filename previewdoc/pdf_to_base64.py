# -*- coding: utf-8 -*-
"""
Created on Tue May 19 13:21:58 2020

@author: Bhanu joshi
"""

import base64 

def convert_pdf_to_base64(filename):
    pdf = open(filename, 'rb') 
    pdf_read = pdf.read() 
    pdf_64_encode = base64.encodestring(pdf_read) 
    return pdf_64_encode