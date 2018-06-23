# Python-Modules
This contents several python modules which can be the resources in the future

# Monitor the inbox and auto forward particular emails.
  This module uses three library to achieve receive email and send the email.
  They are imaplib(receive email), smtplib(send email) and email(get and set the content of email)

  ### hide the content that user enters.
  ```ruby
  getpass.getpass("input") 
  ```
  ### Auto run when boot the computer and keep it running 
  The reason I have a variable named stop is that I want this script can run forever (as long as the computer is booted.) And I don't want to set it as windows service. So I use such little tricky. And I set a .bat file in C:\Users\'username'\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
  and the content of that file is:
  ```ruby
  @echo off
  python c:\auto_forward_email.py %*
  pause  
  ```
  if you don't want a command prompt show when boot the computer, then delete pause ,and directly set the password in script.

  ### search the inbox using special subject.
  ```ruby
  result, data = mail.uid('search', None, 'UNSEEN','HEADER Subject "SUBJECT THAT YOU WANT TO MONITOR"')
  ```

  parse the raw data from imaplib is annoying. A good way is using email library to get the subject, sender and body.

  ### get readable format from raw email:
  ```ruby
  latest_email_uid = data[0].split()[x] # unique ids wrt label selected
  result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')
  //fetch the email body (RFC822) for the given ID
  raw_email = email_data[0][1]
  raw_email_string = raw_email.decode('utf-8')
  //converts byte literal to string removing b''
  email_message = email.message_from_string(raw_email_string)
  ```

  among above:
  ```ruby
  raw_email_string = raw_email.decode('utf-8')
  ```
  This command is important to get string from bytes. Can't directly use str() to transfer.

  ### get subject, sender and body from email.
  Use .walk() and get_payload() to get the body of email
  ```ruby
  subject = email_message['subject']
  sender = email_message['from']
  for part in email_message.walk():
      if part.get_content_type() == "text/plain": # ignore attachments/html
          body = part.get_payload(decode=True)
  ```

  ### Use MIMEMultipart to build the new email that I want to send out.
  ```ruby
  msg = MIMEMultipart()
  msg['From'] = monitor_email
  msg['To'] = forward2email
  msg['Subject'] = 'ANY SUBJECT THAT YOU WANT TO USE'
  message = body.decode('utf-8')
  message = message + sender
  msg.attach(MIMEText(message))
  ```
  ### At last, send the email:
  ```ruby
  smtpObj.send_message(msg)
  ```
