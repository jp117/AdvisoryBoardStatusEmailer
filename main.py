from __future__ import print_function
import httplib2
import os
import datetime
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import send_email
import ABSpreadsheet
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import mimetypes

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

import auth
def get_labels():
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            print(label['name'])

SCOPES = 'https://mail.google.com/'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'
authInst = auth.auth(SCOPES,CLIENT_SECRET_FILE,APPLICATION_NAME)
credentials = authInst.get_credentials()

http = credentials.authorize(httplib2.Http())
service = discovery.build('gmail', 'v1', http=http)

sendDate = datetime.datetime.now().strftime('%m/%d/%Y')

sendInst = send_email.send_email(service)
sender = "john@atlasswitch.com"
recepientList = "gina@atlasswitch.com; harris@atlasswitch.com; matthew@atlasswitch.com; steve@atlasswitch.com; paul@atlasswitch.com; fred@atlasswitch.com; javier@atlasswitch.com; billy@atlasswitch.com; chad@atlasswitch.com; john@atlasswitch.com" 
emailbody = ABSpreadsheet.EmailPendingSubBody()#Logic to loop through spreadsheet and make email

#assembling the email to send
message = sendInst.create_message(sender,recepientList, subject, emailbody)
sendInst.send_message('me',message)
