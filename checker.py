import sys
import os
import imaplib
import getpass
import email
import email.header
import datetime
import xls_parser
from config import *

def process_mailbox(M):
    """
    Do something with emails messages in the folder.
    For the sake of this example, print some headers.
    """

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return

    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return

        msg = email.message_from_bytes(data[0][1])
        hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
        subject = str(hdr)
        print('Message %s: %s' % (num, subject))
        print('Raw Date:', msg['Date'])
        # Now convert to local date-time
        date_tuple = email.utils.parsedate_tz(msg['Date'])
        if date_tuple:
            local_date = datetime.datetime.fromtimestamp(
                email.utils.mktime_tz(date_tuple))
            print ("Local Date:", \
                local_date.strftime("%a, %d %b %Y %H:%M:%S"))

        # save attach
        for part in msg.walk():
            print(part)
            print(part.get_content_type())

            if part.get_content_maintype() == 'multipart':
                continue

            if part.get('Content-Disposition') is None:
                print("no content dispo")
                continue

            # filename = part.get_filename()
            # if not(filename): filename = "test.txt"
            # print(filename)

            # fname = os.path.join('D:\python\mail_ckecker', 'attach{}.xls'.format(datetime.datetime.now()))
            # fp = open(fname, 'wb')
            # fp.write(part.get_payload(decode=1))
            # fp.close

            attach = part.get_payload(decode=1)
            xls_parser.parse(attach)



M = imaplib.IMAP4_SSL('imap.gmail.com')

try:
    # rv, data = M.login(EMAIL_ACCOUNT, getpass.getpass())
    rv, data = M.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
except imaplib.IMAP4.error:
    print ("LOGIN FAILED!!! ")
    sys.exit(1)

print(rv, data)

rv, mailboxes = M.list()
if rv == 'OK':
    print("Mailboxes:")
    print(mailboxes)

rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print("Processing mailbox...\n")
    process_mailbox(M)
    M.close()
else:
    print("ERROR: Unable to open mailbox ", rv)

M.logout()
