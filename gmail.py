import os
import sys
import imaplib
import getpass
import email
import email.header
import datetime
import errno
import mimetypes
import time
import json

EMAIL_ACCOUNT = "patilanup246@gmail.com"
PASSWORD ="Upwork@123"
# Use 'INBOX' to read inbox.  Note that whatever folder is specified,
# after successfully running this script all emails in that folder
# will be marked as read.
EMAIL_FOLDER = "INBOX"
data = {}


def process_mailbox(M):
    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    for num in data[0].split():
        time.sleep(2)
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return

        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
        subject = str(hdr)

        print('Subject %s: %s' % (num, subject))
        print('Raw Date:', msg['Date'])
        print('From :', msg['From'])

        counter = 1
        for part in msg.walk():
            # multipart/* are just containers
            if part.get_content_maintype() == 'multipart':
                continue
            # Applications should really sanitize the given filename so that an
            # email message can't be used to overwrite important files
            filename = part.get_filename()
            if not filename:
                ext = mimetypes.guess_extension(part.get_content_type())
                if not ext:
                    # Use a generic bag-of-bits extension
                    ext = '.bin'
                filename = 'part-%03d%s' % (counter, ext)
            counter += 1
            OUTPUT_DIRECTORY = 'D:/Email'
            f = open('%s/%s' % (OUTPUT_DIRECTORY, filename), 'wb')
            f.write(part.get_payload(decode=True))
            f.close()
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

    with open('data.txt', 'w') as outfile:
        json.dump(data, outfile)


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
