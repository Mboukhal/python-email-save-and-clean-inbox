#!/usr/bin/python3

import time
import os
#import subprocess
import email, imaplib
from email.mime import multipart

IMAP_SERVER = 'imap.gmail.com'
IMAP_PORT = 993
EMAIL_FOLDER = 'INBOX'
OUTPUT_DIRECTORY = 'save_mail'
url_gmail_activation = 'https://support.google.com/accounts/answer/185833'


# def get_email(msg):
#     source = msg['from']
#     to = msg['to']
#     subject = msg['subject']
#     att = msg['attach']
#     count = 1
#     for part in msg.walk():
#         if part.get_content_maintype() == multipart:
#             continue
#         filename = part.get_filename()
#         if not filename:
#             ext = '.html'
#             filename = 'msg-part%80d%s' %(count, ext)
#         count += 1
#     ct = part.get_content_type()

def remove_all(M, nb):
    M.store(nb, '+FLAGS', '\\Deleted')
    M.expunge()

def process_mailbox(M):

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        return
    for num in data[0].split():
        num1 = num
        rv, data = M.fetch(num, '(RFC822)')
        eml = data[0][1].decode('utf-8')
        email_msg = email.message_from_string(eml)
        x = email_msg['date'].split(' ')
        x1 = x[1]+ " "+ x[2] + " " +x[3] + " " + email_msg['from'].split(' ')[2][1:-1]
        path = os.path.join(OUTPUT_DIRECTORY, x1)
        # make save folder
        if not os.path.isdir(OUTPUT_DIRECTORY):
            os.mkdir(OUTPUT_DIRECTORY, mode=0o777)
        if not os.path.isdir(path):
            os.mkdir(path, mode=0o777)
        print(path)
        num = email_msg['from'].split(' ')[2][1:-1]
        if rv != 'OK':
            return
        # unix:
        file_name = '%s/%s' %(path, num)

        # windows:
        #file_name = '%s\\%s' %(path, num)

        # save email as .eml format 
        # you can use this to convert to pdf
        # in folder attachment in 'pdf_converter' 'ecc.exe' 'emailconverter.jar' 'wkhtmltopdf.exe' need to install 'jre-8u333'
        # windows:
        # subprocess.call(f'pdf_converter\ecc.exe "{file_name}.eml"', shell=True)
        # subprocess.call(f'del "{file_name}.eml"', shell=True)

        f = open(file_name + '.eml' , 'wb')
        f.write(data[0][1])
        f.close()
        print(path)

###########################################################
        # save attachment
###########################################################
        
        for part in email_msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            d_file = part.get_filename()
            print(d_file)
            if bool(d_file):
                pfile = os.path.join(path, d_file)
                if not os.path.isfile(pfile):
                    fp = open(pfile, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
###########################################################
        remove_all(M, num1)



def main(gmail, password):
    if gmail[-10:] != '@gmail.com' or len(password) < 8 :
        print('Error password or gmail!')
        return
    while 1:
        M = imaplib.IMAP4_SSL(host=IMAP_SERVER, port=IMAP_PORT)
        try:
            M.login(gmail, password)
        except Exception as msg:
            if str(msg).split(' ')[4] == url_gmail_activation:
                print('Active gmail securty need to be activated in link: ', end='')
                print(url_gmail_activation)
            else:
                print('None valid password or gmail!')
            return
        rv, data = M.select(EMAIL_FOLDER)
        if rv == 'OK':
            print('process_mailbox ok')
            process_mailbox(M)
            M.close()
        M.logout()
        # sleep for 30 min after next cheak for avalbel mails
        time.sleep(1800*60)
        

# if __name__ == '__main__':
    #main('gmail', 'password')