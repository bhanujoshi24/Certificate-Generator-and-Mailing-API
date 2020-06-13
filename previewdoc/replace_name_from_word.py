# -*- coding: utf-8 -*-
"""
Created on Tue May 19 12:36:23 2020

@author: Dharmik joshi
"""
from docx import Document

def replaceNameInWord(filename, participantName):
    doc = Document(filename)
    print(doc)
    for p in doc.paragraphs:
        print(p.text)
        if 'Name'or'name'or'NAME' in p.text:
            print(p.text)
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if 'Name' in inline[i].text:
                    print(inline[i].text)
                    text = inline[i].text.replace('Name', participantName)
                    inline[i].text = text
    doc.save(filename)
    return filename
    