# -*- coding: utf-8 -*-
"""
Created on Tue May 19 12:32:55 2020

@author: Bhanu joshi
"""

import base64

def convertBase64ToWord(attachmentStream, attachmentName):
    attachmentStream = attachmentStream.split(",")[1]
    attachment_bytes = attachmentStream.encode('utf-8')
    with open(attachmentName+'.docx', 'wb') as file_to_save:
        decoded_image_data = base64.decodebytes(attachment_bytes)
        file_to_save.write(decoded_image_data)
    return attachmentName+'.docx'