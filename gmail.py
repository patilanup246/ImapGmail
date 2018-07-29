import datetime
import email
import email.header
import imaplib
import json
import mimetypes
import sys
import time
import uuid

EMAIL_ACCOUNT = "test@gmail.com"  # Gmail Account Username
PASSWORD = "test"  # Gmail Account Password
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
            if counter == 100:
                break
            counter += 1
            time.sleep(1)
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
                    pass  # or you could use 'continue'
            self = {}
            self['Subject'] = subject
            self['Sent Date'] = msg['Date']
            if local_date != '':
                self['Sent Local Date'] = local_date.strftime("%a, %d %b %Y %H:%M:%S")
            self['Name'] = msg['From'].split('<')[0]
            self['From'] = msg['From'].split('<')[1].replace(">", " ")
            self['Attachment File'] = attachmenturl
            self['Email Body'] = emailbody
            self['Bcc'] = msg['Bcc']
            self['Cc'] = msg['Cc']
            self['To'] = msg['To']
            emails.append(self)

            print('Subject ', subject)
            print('Raw Date:', msg['Date'])
            print('From :', msg['From'].split('<')[0])
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
