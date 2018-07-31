import datetime
import email
import email.header
import imaplib
import mimetypes
import socket
import sys
import time
import uuid
from time import gmtime, strftime

from pymongo import MongoClient

Imap_Server = input("Enter Imap_Server [Ex. imap.gmail.com] : ")
EMAIL_ACCOUNT = input("Enter EmailId [Ex.test@gmail.com] : ")  # Gmail Account Username
PASSWORD = input("Enter Password [Ex.password@123] : ")  # Gmail Account Password
Attachment_DIRECTORY = 'D:/Email'  # Attachment file store path
EMAIL_FOLDER = "INBOX"  # Gmail scrap folder name
host_name = socket.gethostname()
host_ip = socket.gethostbyname(host_name)


def process_mailbox(M):
    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    counter = 0
    for num in data[0].split():
        try:
            # time.sleep(1)
            rv, data = M.fetch(num, '(RFC822)')
            if rv != 'OK':
                print("ERROR getting message", num)
                return

            msg = email.message_from_bytes(data[0][1])
            hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
            subject = str(hdr)
            date_tuple = email.utils.parsedate_tz(msg['Date'])
            local_date = ''
            if date_tuple:
                local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
            conn = ''
            try:
                conn = MongoClient('localhost', 27017)
                print("Connected successfully!!!")
            except:
                print("Could not connect to MongoDB")

            # database
            db = conn.hash
            # Created or Switched to collection names: testGmailAnup
            collection = db.testgmail

            if db.testgmail.find({'email_timestamp': str(msg['Date'])}).count() > 0:
                continue

            # Attachment
            attachmenturl = ''
            emailbody = ''
            for part in msg.walk():
                try:
                    # multipart/* are just containers
                    if part.get_content_maintype() == 'multipart':
                        for part in msg.walk():
                            if part.get_content_type() == 'text/plain':
                                emailbody = part.get_payload()  # prints the raw text

                    filename = part.get_filename()
                    if not filename:
                        ext = mimetypes.guess_extension(part.get_content_type())
                        if not ext:
                            ext = '.bin'
                        filename = 'part-%03d%s' % (counter, ext)

                    filename = str(uuid.uuid1()) + filename
                    f = open('%s/%s' % (Attachment_DIRECTORY, filename), 'wb')
                    f.write(part.get_payload(decode=True))
                    f.close()

                    if attachmenturl == '':
                        attachmenturl = Attachment_DIRECTORY + "/" + filename
                    else:
                        attachmenturl = attachmenturl + "," + Attachment_DIRECTORY + "/" + filename
                except Exception:
                    continue  # or you could use 'continue'

            timestamp = int(time.time())  # timestamp

            # From if not found
            email_sender = ''
            email_sender_id = ''
            try:
                if '<' in msg['From']:
                    email_sender = msg['From'].split('<')[0]
                    email_sender_id = msg['From'].split('<')[1].replace(">", " ")
                else:
                    email_sender_id = msg['From']
            except:
                print("")

            # To if not found
            email_recipeint = ''
            email_recipient_id = ''
            try:
                if '<' in msg['To']:
                    email_recipeint = msg['To'].split('<')[0]
                    email_recipient_id = msg['To'].split('<')[1].replace(">", " ")
                else:
                    email_recipient_id = msg['To']
            except:
                print("")

            # CC if not found
            email_recipeint_CC = ''
            email_recipient_CC_id = ''
            try:
                if '<' in msg['Cc']:
                    email_recipeint_CC = msg['Cc'].split('<')[0]
                    email_recipient_CC_id = msg['Cc'].split('<')[1].replace(">", " ")
                else:
                    email_recipient_CC_id = msg['Cc']
            except:
                print("")

            # CC if not found
            email_recipeint_CCO = ''
            email_recipient_CCO_ID = ''
            try:
                if '<' in msg['Bcc']:
                    email_recipeint_CCO = msg['Bcc'].split('<')[0]
                    email_recipient_CCO_ID = msg['Bcc'].split('<')[1].replace(">", " ")
                else:
                    email_recipient_CCO_ID = msg['Bcc']
            except:
                print("")


            emp_rec1 = {
                "tib": "8f7074d8-a520-4f7d-b2d3-09dc36acb5fd",
                "tib_name": "TPEMAIL",
                "mail_box_name": "Frederico Gmail",
                "id_mail_box": str(uuid.uuid1()),
                "time": timestamp,
                "time_zone": strftime("%z", gmtime()),
                "email_subject": subject,
                "email_sender": email_sender,
                "email_sender_id": email_sender_id,
                "email_recipeint": email_recipeint,
                "email_recipient_id": email_recipient_id,
                "email_timestamp": msg['Date'],
                "email_header": "",
                "email_body": emailbody,
                "email_seq": "",
                "email_text_content": "",
                "email_html_content": "",
                "email_eml_content": "",
                "email_links": "",
                "email_images": attachmenturl,
                "email_template_id": "",
                "email_track_link": "",
                "email_recipeint_CC": email_recipeint_CC,
                "email_recipient_CC_id": email_recipient_CC_id,
                "email_recipeint_CCO": email_recipeint_CCO,
                "email_recipient_CCO_ID": email_recipient_CCO_ID,
                "hash_owner_id": "g00zNU6n7WfhUI1u4A5ebxSN0732",
                "hash_sender_id": "",
                "hash_recipt_id": "",
                "hash_sender_name": "",
                "hash_receipt_name": "",
                "hash_recipt_CC_id": "",
                "hash_recipt_CCO_ID": "",
                "email_atach": "",
                "hub_group_id": "da0a7b22-fb15-46e0-9f5a-019263d79e36",
                "data_sinc": "",
                "action": "",
                "role": "",
                "layout_role": "",
                "group_id": "9529b03b-38e8-4bdc-aa62-8055a4c36a55",
                "IP_machine": host_ip
            }
            # Insert Data
            rec_id1 = collection.insert_one(emp_rec1)
            print("Data inserted with record ids", rec_id1)

            counter += 1
            print(str(counter), "]")
            print('Subject :', subject)
            print('Raw Date:', msg['Date'])
            print('From :', msg['From'].split('<')[0])
            print("")
        except Exception as e:
            print('Main Fail: ' + str(e))
            continue  # or you could use 'continue'


M = imaplib.IMAP4_SSL(Imap_Server)

try:
    rv, data = M.login(EMAIL_ACCOUNT, PASSWORD)
except imaplib.IMAP4.error:
    print("LOGIN FAILED!!! ")
    sys.exit(1)

print(rv, data)

rv, mailboxes = M.list()
if rv == 'OK':
    print("Mailboxes:")

rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print("Processing mailbox...\n")
    process_mailbox(M)
    M.close()
else:
    print("ERROR: Unable to open mailbox ", rv)

M.logout()
