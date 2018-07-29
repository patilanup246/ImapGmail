import datetime
import email
import email.header
import imaplib
import json
import mimetypes
import sys
import time
import uuid

EMAIL_ACCOUNT = "patilanup246@gmail.com"
PASSWORD ="Upwork@123"
OUTPUT_DIRECTORY = 'D:/Email'
EMAIL_FOLDER = "INBOX"
self = {}
Emails = []

def process_mailbox(M):
    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    for num in data[0].split():
        try:
            time.sleep(2)
            rv, data = M.fetch(num, '(RFC822)')
            if rv != 'OK':
                print("ERROR getting message", num)
                return

            msg = email.message_from_bytes(data[0][1])
            hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
            subject = str(hdr)

            self['Subject'] = subject
            self['Raw Date'] = msg['Date']
            self['Name From'] = msg['From'].split('<')[0]
            self['Email From'] = msg['From'].split('<')[1]
            Emails.append(self)

            print('Subject %s: %s' % (num, subject))
            print('Raw Date:', msg['Date'])
            print('From :', msg['From'].split('<')[0])
            print('email from :', msg['From'].split('<')[1].replace(">", " "))
            counter = 1
            for part in msg.walk():
                try:
                    # multipart/* are just containers
                    if part.get_content_maintype() == 'multipart':
                        continue
                    filename = part.get_filename()
                    if not filename:
                        ext = mimetypes.guess_extension(part.get_content_type())
                        if not ext:
                            ext = '.bin'
                        filename = 'part-%03d%s' % (counter, ext)
                    counter += 1

                    f = open('%s/%s' % (OUTPUT_DIRECTORY, str(uuid.uuid1()) + filename), 'wb')
                    f.write(part.get_payload(decode=True))
                    f.close()
                except Exception:
                    pass  # or you could use 'continue'
            # if msg.is_multipart():
            #     for part in msg.walk():
            #         # each part is a either non-multipart, or another multipart message
            #         # that contains further parts... Message is organized like a tree
            #         if part.get_content_type() == 'text/plain':
            #             print(part.get_payload())  # prints the raw text
            #         elif part.get_content_type() == 'application/msword':
            #             name = part.get_param('name') or 'MyDoc.doc'
            #             f = open(name, 'wb')
            #             f.write(part.get_payload(None, True))
            #             f.close()
            # else:
            #     print(msg.get_payload())
            # Now convert to local date-time
            date_tuple = email.utils.parsedate_tz(msg['Date'])
            if date_tuple:
                local_date = datetime.datetime.fromtimestamp(
                    email.utils.mktime_tz(date_tuple))
                print("Local Date:", local_date.strftime("%a, %d %b %Y %H:%M:%S"))
        except Exception:
            pass  # or you could use 'continue'

    with open("test.json", 'w') as f:  # you don't need f.close() with a file manager
        json.dump(data, f)


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
