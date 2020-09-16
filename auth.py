import base64
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
import mimetypes
import pickle
import os
from apiclient import errors
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import re
import email
import base64

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://mail.google.com/','https://googleapis.com/']

def get_msg(service,msg_id):
    try:
        message_list= service.users().messages().get(userId='me',id=msg_id,format='raw').execute()
        msg_raw=base64.urlsafe_b64decode(message_list['raw'].encode('ASCII'))
        #print(message_list) 
        msg_str=email.message_from_bytes(msg_raw)
        #print(msg_str)
        content_types=msg_str.get_content_maintype()
        #print(content_types)
        if content_types=='multipart':
            part1,part2=msg_str.get_payload()
            return {'to':msg_str['Return-Path'],'message':part1.get_payload(),'subject':msg_str['Subject'],'thread_id':message_list['threadId'],'messageid':msg_str['Message-ID']}
        else:
            return {'to':msg_str['Return-Path'],'message':part1.get_payload(),'subject':msg_str['Subject'],'thread_id':message_list['threadId'],'messageid':msg_str['Message-ID']}

    except(errors.HttpError, error):
        print("An error occured:")


def search_msgs(service,user_id,search_string):
    try:
        search_id=service.users().messages().list(userId=user_id,q=search_string).execute()
        num_results=search_id['resultSizeEstimate']
        final_list=[]
        if(num_results!=0):
            msg_id=search_id['messages'] 
            for i in msg_id:
                final_list.append(i['id'])
            return final_list
        else:
            print("no results found for the string")
            return ""

    except(errors.HttpError, error):
        print("An error occured:")


def get_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    return service

service=get_service()
id=search_msgs(service,'me','label:unread')
res=[]
for i in id:
    res.append(get_msg(service,i))
print(res[0]['message'])

#product=['[Pp]roduct+' '*+'-'*+[Aa]','[Pp]roduct+' '*+'-'*+[Bb]','[Pp]roduct+' '*+'-'*+[Cc]']
products=['product A','product B','product C']
for i in res:
    for j in products:
        #print(j,i[1])
        if re.search(j,i['message']):
            i['reply']=j
            break
print(res)

def send_message(sender,to,subject,message_text,threadId,mId):
    emailMsg = message_text
    mimeMessage = MIMEMultipart()
    mimeMessage['In-Reply-To']=mId
    mimeMessage['References']=mId
    mimeMessage['to'] = to
    mimeMessage['subject'] = subject
    mimeMessage.attach(MIMEText(emailMsg, 'plain'))
    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
    message = service.users().messages().send(userId='me', body={'raw': raw_string,'threadId':threadId}).execute()
    return message

service=get_service()
user_id='me'
to=res[0]['to']
mid=res[0]['messageid']
subject=res[0]['subject']
print(to,subject,mid)
msg=send_message('zcompanybot@gmail.com',to,subject,'Product A is what you asked',res[0]['thread_id'],mid)
print(msg)
