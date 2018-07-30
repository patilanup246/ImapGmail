import datetime
import email
import email.header
import imaplib
import json
import mimetypes
import sys
import uuid

EMAIL_ACCOUNT = "fred.medeiros72@gmail.com"  # Gmail Account Username
PASSWORD = "fjmedeiros"  # Gmail Account Password
Attachment_DIRECTORY = 'D:/Email'  # Attachment file store path
EMAIL_FOLDER = "INBOX"  # Gmail scrap folder name
FinalOutput_DIRECTORY = 'E:/Python/ImapGmail/Output/'  # Json final output file store path


def process_mailbox(M):
    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    emails = []
    counter = 0
    for num in data[0].split():
        try:
            # if counter == 100:
            #     break

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
                local_date = datetime.datetime.fromtimestamp(
                    email.utils.mktime_tz(date_tuple))

            attachmenturl = ''
            emailbody = ''
            for part in msg.walk():
                try:
                    # multipart/* are just containers
                    if part.get_content_maintype() == 'multipart':
                        for part in msg.walk():
                            # each part is a either non-multipart, or another multipart message
                            # that contains further parts... Message is organized like a tree
                            if part.get_content_type() == 'text/plain':
                                emailbody = part.get_payload()  # prints the raw text

                    filename = part.get_filename()
                    if not filename:
                        ext = mimetypes.guess_extension(part.get_content_type())
                        if not ext:
                            pass
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
                    pass  # or you could use 'continue'
            self = {}

            self['Subject'] = subject
            self['Sent Date'] = msg['Date']
            self['Sent Local Date'] = local_date.strftime("%a, %d %b %Y %H:%M:%S")
            self['Name'] = msg['From'].split('<')[0]
            self['From'] = msg['From'].split('<')[1].replace(">", " ")
            self['Attachment File'] = attachmenturl
            self['Email Body'] = emailbody
            self['Bcc'] = msg['Bcc']
            self['Cc'] = msg['Cc']
            self['To'] = msg['To']
            emails.append(self)

            # try:
            #     conn = MongoClient('localhost', 27017)
            #     print("Connected successfully!!!")
            # except:
            #     print("Could not connect to MongoDB")
            # # database
            # db = conn.hash
            # # Created or Switched to collection names: ib
            # collection = db.ib
            # emp_rec1 = {
            #     "tib": "8f7074d8-a520-4f7d-b2d3-09dc36acb5fd",
            #     "tib_name": "TPEMAIL",
            #     "mail_box_name": "Frederico Gmail",
            #     "id_mail_box": "use here when generate to UUID v4 - MONGO",
            #     "time": msg['Date'],
            #     "time_zone": local_date.strftime("%a, %d %b %Y %H:%M:%S"),
            #     "email_subject": subject,
            #     "email_sender": msg['From'].split('<')[0],
            #     "email_sender_id": msg['From'].split('<')[1].replace(">", " "),
            #     "email_recipeint": "",
            #     "email_recipient_id": "",
            #     "email_timestamp": "",
            #     "email_header": "",
            #     "email_body": emailbody,
            #     "email_seq": "",
            #     "email_text_content": "",
            #     "email_html_content": "",
            #     "email_eml_content": "",
            #     "email_links": attachmenturl,
            #     "email_images": "scrap all imagens inside email",
            #     "email_template_id": "if we sent using any template",
            #     "email_track_link": "if this email had any track link to ansewr",
            #     "email_recipeint_CC": msg['Cc'],
            #     "email_recipient_CC_id": "",
            #     "email_recipeint_CCO": msg['Bcc'],
            #     "email_recipient_CCO_ID": "",
            #     "hash_owner_id": "g00zNU6n7WfhUI1u4A5ebxSN0732",
            #     "hash_sender_id": "",
            #     "hash_recipt_id": "",
            #     "hash_sender_name": "",
            #     "hash_receipt_name": "",
            #     "hash_recipt_CC_id": "",
            #     "hash_recipt_CCO_ID": "",
            #     "email_atach": "arry with all atachments and path saved",
            #     "hub_group_id": "da0a7b22-fb15-46e0-9f5a-019263d79e36",
            #     "data_sinc": "",
            #     "action": "",
            #     "role": "",
            #     "layout_role": "",
            #     "group_id": "9529b03b-38e8-4bdc-aa62-8055a4c36a55",
            #     "IP_machine": "ip_machine_request"
            # }
            # # Insert Data
            # rec_id1 = collection.insert_one(emp_rec1)
            #
            # print("Data inserted with record ids", rec_id1)

            counter += 1
            print(str(counter), "]")
            print('Subject :', subject)
            print('Raw Date:', msg['Date'])
            print('From :', msg['From'].split('<')[0])
            print("")
        except Exception:
            pass  # or you could use 'continue'

    with open(FinalOutput_DIRECTORY + "emails" + str(uuid.uuid1()) + ".json",
              'w') as f:  # you don't need f.close() with a file manager
        json.dump(emails, f, indent=4, sort_keys=True)


M = imaplib.IMAP4_SSL('imap.gmail.com')

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
