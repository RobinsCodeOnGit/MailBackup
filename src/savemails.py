from email.message import EmailMessage
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import getpass
import base64
import json

resources = json.load(open('resources.json'))

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

def encode(s:str):
    return base64.b64encode(s.encode(encoding='utf-8')).decode(encoding='utf-8')

def decode(s:str):
    return base64.b64decode(s).decode(encoding='utf-8')

# create an IMAP4 class with SSL 
imap = imaplib.IMAP4_SSL(resources['account']['server'])
# authenticate
imap.login(resources['account']['username'], decode(resources['account']['password_enc']))

folders = [ x.decode().split('"/" ')[1] for x in imap.list()[1]]

for folder in folders:
    status, messages = imap.select(folder)
    if status != 'OK':
        print("Errro accessing folder: " + folder)
    
    messages = int(messages[0])

no_of_messages = int(imap.select('INBOX')[1][0])

res, msg = imap.fetch(str(no_of_messages), "(RFC822)")

mail = email.message_from_bytes(msg[0][1])
for key in list(dict.fromkeys(mail.keys())):
    print(key + ' ##############################')
    for i,j in enumerate(mail.get_all(key)):
        print('############## ' + str(i) + ' ###############')
        print(j)