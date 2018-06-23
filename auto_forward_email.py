
##this script can check and send unread email with specail subject as long as it's running !!!!!
##But it can only work when it runs: On windows, set this script as windows service.
##                                   On Linux, just set this script as background (use &)
import imaplib
import smtplib
import email
import getpass
import time
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

USER_NAME = getpass.getuser()


##########################################################
################# Variables ##############################
##########################################################
start_time = time.time()
period = 300           #run 300 seconds to test whether this script can monitor the inbox while running.
monitor_email = "name@mail.com"
forward2email = "another_name@mail.com"
password = getpass.getpass("Please input password: ")
stop = 'a'

#inistial the server.
#imaplib monitor the inbox and receive the email.
#smtplib sned the email.
mail = imaplib.IMAP4_SSL('imap.gmail.com')
smtpObj = smtplib.SMTP('smtp.gmail.com',587)
smtpObj.ehlo()
smtpObj.starttls()
# imaplib module implements connection based on IMAPv4 protocol

mail.login(monitor_email, password)
smtpObj.login(monitor_email, password)

mail.list()
mail.select('inbox')

##########################################################
#### search unread email according to subject#############
#### then forward it to another email ####################
##########################################################
#force the script run forever(better way to do that is adding this script to windows service.)
while(stop != 's'):
    #serach and get uid of unread mail with special subject.
    result, data = mail.uid('search', None, 'UNSEEN','HEADER Subject "SUBJECT THAT YOU WANT TO MONITOR"')
    # search and return uids instead
    i = len(data[0].split()) # data[0] is a space separate string
    for x in range(i):
        latest_email_uid = data[0].split()[x] # unique ids wrt label selected
        print("this is: ",latest_email_uid)
        result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')
        # fetch the email body (RFC822) for the given ID
        raw_email = email_data[0][1]

        #continue inside the same for loop as above
        raw_email_string = raw_email.decode('utf-8')
        # converts byte literal to string removing b''
        email_message = email.message_from_string(raw_email_string)
        # this will loop through all the available multiparts in mail
        subject = email_message['subject']
        sender = email_message['from']
        print('From: ',sender)
        print('subject: ',subject)
        for part in email_message.walk():
            if part.get_content_type() == "text/plain": # ignore attachments/html
                body = part.get_payload(decode=True)
        print("body is: ",body.decode('utf-8'))
        msg = MIMEMultipart()
        msg['From'] = monitor_email
        msg['To'] = forward2email
        msg['Subject'] = 'ANY SUBJECT THAT YOU WANT TO USE'
        message = body.decode('utf-8')
        message = message + sender     
        msg.attach(MIMEText(message))
        smtpObj.send_message(msg)
        print("finish sending the unread email and mark it as read\n")
