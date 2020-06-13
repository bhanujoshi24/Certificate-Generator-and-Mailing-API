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
#
import os
# import base64_to_word #base64_to_word will be used to convert base64 String into a docx file and will save it on the system
# import replace_name_from_word #replace_name_from_word will be used to replace a name from word to participant name and will save it on the system
# import docx_to_pdf #after the success full replacement of name will convert the docx file to pdf and save to system
# import pdf_to_base64 #this module is use to convert pdf into base64 string will be used for preview
import requests
from . import base64_to_word,docx_to_pdf,replace_name_from_word

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
# Create your views here.

@api_view(["POST"])
def MailAPI(jsonobj):
    try:
        data=jsonobj.data
        print(jsonobj.data)
        print(type(data))
        senderEmailAddress=data['senderEmailAddress']
        senderEmailPassword=data['senderEmailPassword']
        receiverName=data['receiverName']
        receiverEmailAddress=data['receiverEmailAddress']
        emailBody=data['emailBody']
        id=data['id']
        emailSubject=data['emailSubject']
        # emailattachmentname=data['emailattachmentname']
        emailattachment=data['emailattachment']
        try:
            print("Bhanu is here 31")
            emailAttachment=emailattachment
            # first step is to convert the base64 string to docx file save it with the receiver name ex: dharmik_joshi.docx
            docx_file_name = base64_to_word.convertBase64ToWord(emailAttachment, str(receiverName).replace(" ", "_"))
            print("Bhanu is here 32")

            # second step is to put the name of receiver in certificate
            certificate_of_participant_docx = replace_name_from_word.replaceNameInWord(docx_file_name, receiverName)
            print("Bhanu is here 33")

            # third step is to convert the docx file into PDF
            print(type(certificate_of_participant_docx))

            certificate_of_participant_pdf = docx_to_pdf.convertFile(certificate_of_participant_docx)
            print("Bhanu is here 34")

            #forth step is to mail a participant

            fromaddr = senderEmailAddress
            toaddr = receiverEmailAddress

            # instance of MIMEMultipart
            msg = MIMEMultipart()

            # storing the senders email address
            msg['From'] = senderEmailAddress

            # storing the receivers email address
            msg['To'] = receiverEmailAddress

            # storing the subject
            msg['Subject'] = emailSubject

            # string to store the body of the mail
            body = emailBody

            # attach the body with the msg instance
            msg.attach(MIMEText(body, 'plain'))

            # open the file to be sent
            filename = certificate_of_participant_pdf
            attachment = open(certificate_of_participant_pdf, "rb")

            # instance of MIMEBase and named as p
            p = MIMEBase('application', 'octet-stream')

            # To change the payload into encoded form
            p.set_payload((attachment).read())

            # encode into base64
            encoders.encode_base64(p)

            p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

            # attach the instance 'p' to instance 'msg'
            msg.attach(p)

            # creates SMTP session
            s = smtplib.SMTP('smtp.gmail.com', 587)

            # start TLS for security
            s.starttls()

            # Authentication
            s.login(fromaddr, senderEmailPassword)

            # Converts the Multipart msg into a string
            text = msg.as_string()

            # sending the mail
            s.sendmail(fromaddr, toaddr, text)

            # terminating the session
            s.quit()

            # print("Ended")

            #final step step is to remove the files after the mail this will happen irrespective wheather the mail is sent or not
            # if os.path.exists(docx_file_name):
            #     os.remove(docx_file_name)
            #
            # if os.path.exists(certificate_of_participant_pdf):
            #     os.remove(certificate_of_participant_pdf)

        except Exception as e:
            return JsonResponse(e, safe=False)
        return  JsonResponse("Successfully sent", safe=False)
    except Exception as e:
        return Response(e.args[0],status.HTTP_400_BAD_REQUEST)