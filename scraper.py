import imaplib, email,os

imap_host = 'imap.gmail.com'
user = 'raghavjuyal9653@gmail.com'
password = 'nvdkifremaybnqai'


def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)
        


# connect to host using SSL
imap = imaplib.IMAP4_SSL(imap_host)

## login to server
imap.login(user, password)
imap.select("INBOX")


# result,data = imap.fetch(b'37','(RFC822)')
# raw = email.message_from_bytes(data[0][1])

# res, messages = imap.select('"[Gmail]/Sent Mail"')  

res, messages = imap.select('INBOX') 


def attachment(subject):
    type, data = imap.search(None, 'ALL')
    for num in data[0].split():
        typ, data = imap.fetch(num, '(RFC822)' )
        raw_email = data[0][1]
    # converts byte literal to string removing b''
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)
    # downloading attachments
        for part in email_message.walk():
            # this part comes from the snipped I don't understand yet... 
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            if bool(fileName):
                filePath = os.path.join('c:/CODE/Revvy/', fileName)
                if not os.path.isfile(filePath) :
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                print(f'Downloaded "{fileName}" from email titled "{From}"')



# print(data)
data = int(messages[0])


for i in range(data, data - 3, -1):
    res, msg = imap.fetch(str(i), "(RFC822)")     
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
              
            # getting the sender's mail id
            From = str(msg["FROM"])
            check_mail = From.split(' ')
            subject = msg["Subject"]

            for x in check_mail:
                if x == "<raghavraghav456147@gmail.com>" and subject == 'CSV':
                    print('ok')
        
                    # printing the details
                    print("From : ", From)
                    print("subject : ", subject)
                    # downloading attachments
                    attachment(From)
                    