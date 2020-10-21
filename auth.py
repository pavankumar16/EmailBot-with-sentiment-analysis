import base64
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
import mimetypes
import pickle
import os
#from apiclient import errors
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
            return {'to':msg_str['Return-Path'],'message':part1.get_payload(),'subject':msg_str['Subject'],'thread_id':message_list['threadId'],'messageid':msg_str['Message-ID'],'id':msg_id}
        else:
            return {'to':msg_str['Return-Path'],'message':part1.get_payload(),'subject':msg_str['Subject'],'thread_id':message_list['threadId'],'messageid':msg_str['Message-ID'],'id':msg_id}

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
            print(".")
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
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    return service

def send_message(sender,to,subject,message_text,threadId,mId,attached_file=None):
    emailMsg = message_text
    mimeMessage = MIMEMultipart()
    mimeMessage['In-Reply-To']=mId
    mimeMessage['References']=mId
    mimeMessage['to'] = to
    mimeMessage['subject'] = subject
    mimeMessage.attach(MIMEText(emailMsg, 'plain'))
    if(attached_file != None):
        my_mimetype, encoding = mimetypes.guess_type(attached_file)
        if my_mimetype is None or encoding is not None:
        	my_mimetype = 'application/octet-stream'
        main_type, sub_type = my_mimetype.split('/', 1)
        if main_type == 'text':
            print("text")
            temp = open(attached_file, 'r')  # 'rb' will send this error: 'bytes' object has no attribute 'encode'
            attachment = MIMEText(temp.read(), _subtype=sub_type)
            temp.close()
        elif main_type == 'image':
            print("image")
            temp = open(attached_file, 'rb')
            attachment = MIMEImage(temp.read(), _subtype=sub_type)
            temp.close()
        elif main_type == 'audio':
            print("audio")
            temp = open(attached_file, 'rb')
            attachment = MIMEAudio(temp.read(), _subtype=sub_type)
            temp.close()
        elif main_type == 'application' and sub_type == 'pdf':   
            temp = open(attached_file, 'rb')
            attachment = MIMEApplication(temp.read(), _subtype=sub_type)
            temp.close()
        else:                              
            attachment = MIMEBase(main_type, sub_type)
            temp = open(attached_file, 'rb')
            attachment.set_payload(temp.read())
            temp.close()
        filename = os.path.basename(attached_file)
        attachment.add_header('Content-Disposition', 'attachment', filename=filename) # name preview in email
        mimeMessage.attach(attachment)
    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
    message = service.users().messages().send(userId='me', body={'raw': raw_string,'threadId':threadId}).execute()
    return message

file_path='C:\\Users\\pavan\\Desktop\\capstone\\catalogs\\A.pdf'
isExist = os.path.exists(file_path)
print(isExist)

while(True):
    service=get_service()
    id=search_msgs(service,'me','label:unread')
    res=[]
    for i in id:
        res.append(get_msg(service,i))
    #print(res[0]['message'])
    if(len(res)==0):
        continue

    #products=['[Pp]roduct+' '*+'-'*+[Aa]','[Pp]roduct+' '*+'-'*+[Bb]','[Pp]roduct+' '*+'-'*+[Cc]']
    products=['[Pp]roduct A','[Pp]roduct B','[Pp]roduct C']
    for i in res:
        i['reply']=None
        for j in products:
            if re.search(j,i['message']):
                i['reply']=j[-1]
                break
    print(res)

    service=get_service()
    user_id='me'
    for i in range(len(res)):
        to=res[i]['to']
        mid=res[i]['messageid']
        subject=res[i]['subject']
        print(to,subject,mid)
        if(res[i]['reply']!=None):
            file_path='C:\\Users\\pavan\\Desktop\\capstone\\catalogs\\'+str(res[i]['reply'])+'.pdf'
            if(os.path.exists(file_path)):
                service.users().messages().modify(userId='me',id=res[i]['id'],body={'removeLabelIds': ['UNREAD']}).execute()
                msg=send_message('zcompanybot@gmail.com',to,subject,'here is a catalog of product you asked for your reference',res[i]['thread_id'],mid,file_path) #replace reply string with res[i]['reply']
        else:
            continue
        print(msg)
