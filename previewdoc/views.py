from django.shortcuts import render
from django.shortcuts import render
from django.http import Http404
from rest_framework.views import APIView
from rest_framework.decorators import api_view
from rest_framework.response import Response
from rest_framework import status
from django.http import JsonResponse
from django.core import serializers
from django.conf import settings
import json
import requests
from . import base64_to_word,docx_to_pdf,pdf_to_base64,replace_name_from_word
import os
# import base64_to_word #base64_to_word will be used to convert base64 String into a docx file and will save it on the system
# import replace_name_from_word #replace_name_from_word will be used to replace a name from word to participant name and will save it on the system
# import docx_to_pdf #after the success full replacement of name will convert the docx file to pdf and save to system
# import pdf_to_base64 #this module is use to convert pdf into base64 string will be used for preview

# Create your views here.

@api_view(["POST"])
def previewdocFile(jsonobj):

    data = jsonobj.data
    # print(jsonobj.data)
    # print(type(data))
    emailAttachment = data['emailattachment']
    # print(type(emailAttachment))
    try:
        # print("Bhanu is here 1")
        # first step is to convert the base64 string to docx file save it with the receiver name ex: dharmik_joshi.docx
        docx_file_name = base64_to_word.convertBase64ToWord(emailAttachment, "preview")
        print("Bhanu is here 2")
        # second step is to put the name of receiver in certificate
        certificate_of_participant_docx = replace_name_from_word.replaceNameInWord(docx_file_name, "Dharmik Joshi")
        print("Bhanu is here 3 ")
        # third step is to convert the docx file into PDF
        certificate_of_participant_pdf = docx_to_pdf.convertFile(certificate_of_participant_docx)
        print("Bhanu is here 4")
        # forth step is to convert pdf to base64 string
        pdf_in_base64_string = pdf_to_base64.convert_pdf_to_base64(certificate_of_participant_pdf)

        print(type(pdf_in_base64_string))
        print("Bhanu is here")
        # final step step is to remove the files after the conversion
        # if os.path.exists(docx_file_name):
        #     os.remove(docx_file_name)
        #
        # if os.path.exists(certificate_of_participant_pdf):
        #     os.remove(certificate_of_participant_pdf)

        return JsonResponse(str(pdf_in_base64_string), safe=False)
        # return pdf_in_base64_string
    except Exception as e:
        return Response(e.args[0], status.HTTP_400_BAD_REQUEST)