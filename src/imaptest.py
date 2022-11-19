import imaplib
import pprint

IMAP_HOST = 'imap.gmail.com'
IMAP_USER = 'fstemarie.bb@gmail.com'
IMAP_PASS = 'uswmpffhihegkuxm'

with imaplib.IMAP4_SSL(IMAP_HOST) as imap:
    imap.login(IMAP_USER, IMAP_PASS)

    [results, data] = imap.list("I")
    pprint.pprint(results)
    # pprint.pprint(data)
    for directory in data:
        print(directory.decode().split('"/"')[-1])
    # tmp, data = imap.search(None, 'ALL')
    # pprint.pprint(tmp)
    # for num in data[0].split():
    #     tmp, data = imap.fetch(num, '(RFC822)')
    #     print(f'Message: {num}\n')
    #     # pprint.pprint(data[0][1])
    #     break
