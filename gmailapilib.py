import io
from openpyxl import *  
# this will teach me not to do * imports :)
# zbog import * mi ne radi open()
# pa moram io.open()
from datetime import *
import os
import json
import base64
import email

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.message import EmailMessage

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from gpt import process_mail_gpt




SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send"
]


max_results = 90

# aux functions za parse_email_body
def get_service():  # konekcija na server

    def get_creds():
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("token.json"):
          creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
          if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
          else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
          # Save the credentials for the next run
          with io.open("token.json", "w") as token:
            token.write(creds.to_json())

        return creds
    
    creds = get_creds()
    service = build("gmail", "v1", credentials=creds)
    return service

def get_labelId(service,labelName):
    
    if labelName == 'INBOX': return 'INBOX'

    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    for label in labels:
       if label['name'] == labelName:
          return label['id']

def get_charset(part, default="utf-8"):
    if part.get_content_charset():
        return part.get_content_charset()
    if "charset" in part.get("Content-Type", ""):
        return part.get_content_type().split("charset=")[-1].strip()
    return default

def get_mime_message(service, user_id, msg_id):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id, format='raw').execute()
        msg_raw = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
        mime_msg = email.message_from_bytes(msg_raw)
        return mime_msg
    except HttpError as error:
        print('An error occurred: %s' % error)
    
def create_message_with_pdf_attachment(sender, to, subject, message_text, pdf_file_path):
    """Create a message for an email with a PDF attachment."""
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    # Add the message body
    msg = MIMEText(message_text)
    message.attach(msg)

    # Add the PDF attachment
    if not pdf_file_path.lower().endswith('.pdf'):
        raise ValueError("The file must be a PDF.")
    
    with io.open(pdf_file_path, 'rb') as fp:
        part = MIMEBase('application', 'pdf')
        part.set_payload(fp.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_file_path))
        message.attach(part)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw}



def create_pdf_mail_reply(original_message_id, user_id, reply_text, pdf_file_path):
    service = get_service()
    original_message = service.users().messages().get(userId=user_id, id=original_message_id).execute()
    thread_id = original_message['threadId']

    message = MIMEMultipart()
    message['to'] = "aleksandarvasiljevic11@gmail.com"
    message['from'] = user_id
    message['subject'] = ""

    msg = MIMEText(reply_text)
    message.attach(msg)

    if not pdf_file_path.lower().endswith('.pdf'):
        raise ValueError("The file must be a PDF.")
    
    with io.open(pdf_file_path, 'rb') as fp:
        part = MIMEBase('application', 'pdf')
        part.set_payload(fp.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_file_path))
        message.attach(part)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw, 'threadId': thread_id}




def send_message(service, user_id, message):
    """Send an email message."""
    try:
        message = service.users().messages().send(userId=user_id, body=message).execute()
        print(f'Message Id: {message["id"]}')
        return message
    except Exception as error:
        print(f'An error occurred: {error}')
        return None
    

def move_message_to_label(service, user_id, msg_id, remove_label_id, add_label_id):
    """Move a message from one label to another."""
    try:
        message = service.users().messages().modify(
            userId=user_id,
            id=msg_id,
            body={
                'removeLabelIds': [remove_label_id],
                'addLabelIds': [add_label_id]
            }
        ).execute()
        print(f'Message Id: {message["id"]} - Moved successfully')
        return message
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None