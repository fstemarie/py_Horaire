"""
	* pyHoraire *
	Prends mon horaire a partir du email et va remplir mon calendrier Google 
"""

"""
RFC2045 - Multipurpose Internet Mail Extensions (MIME) Part One
    https://tools.ietf.org/html/rfc2045.html

RFC2046 - Multipurpose Internet Mail Extensions (MIME) Part Two
    https://tools.ietf.org/html/rfc2046.html

RFC2047 - MIME (Multipurpose Internet Mail Extensions) Part Three
    https://tools.ietf.org/html/rfc2047.html

RFC2183 - Communicating Presentation Information in Internet Messages
    https://tools.ietf.org/html/rfc2183.html

RFC2231 - MIME Parameter Value and Encoded Word Extensions
    https://tools.ietf.org/html/rfc2231.html
"""


# If modifying these scopes, delete the file token.pickle.


import base64
import os.path
import pickle
import traceback
from pprint import pprint
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from bs4.element import NavigableString, Tag
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
SCOPES = ['https://mail.google.com/']
LBLUNPROCESSED = 'Horaire'
LBLPROCESSED = 'Processed'


def get_service():
    """ Get the credentials for access to the Google API """
    creds = None
    flow = None
    token = None

    # The file token.pickle stores the user's access and refresh tokens,
    # and is created automatically when the authorization flow completes
    # for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    print(creds)
    return build('gmail', 'v1', credentials=creds)


def get_parts(mimetype: str, part: object) -> list:
    parts = []
    if 'payload' in part:
        # If we are at the root of the message tree
        part = part['payload']
    if part['mimeType'].startswith('multipart'):
        for p in part['parts']:
            parts.extend(get_parts(mimetype, p))
    elif part['mimeType'] == mimetype:
        parts.append(part)
    return parts


def extract_date(gmail_msg: dict) -> datetime:
    msg_date = filter(
        lambda h: h['name'] == 'Date',
        gmail_msg['payload']['headers'])
    msg_date = list(msg_date)[0]['value']
    msg_date = msg_date.split()[1:4]
    msg_date = ' '.join(msg_date)
    msg_date = datetime.strptime(msg_date, "%d %b %Y")
    return msg_date


def fix_date_token(cell: str, msg_date: datetime) -> datetime:
    date = None

    date = cell.split()[1:3]
    date.append(str(msg_date.year))
    date = ','.join(date)
    date = datetime.strptime(date, '%b,%d,%Y')
    if date < msg_date:
        # If schedule is for january next year
        date = date.replace(year=date.year + 1)
    return date


def fix_time_token(time_token: str) -> datetime:
    date = None

    if len(time_token) in (4, 5):
        if time_token.find(' ') > -1:
            date = datetime.strptime(time_token, '%I %p').time()
        else:
            date = datetime.strptime(time_token, '%I%p').time()
    elif len(time_token) in (7, 8):
        date = datetime.strptime(time_token, '%I:%M %p').time()
    return date



def process_workday(date: datetime, cell: str) -> list:
    tokens = None
    start = None
    end = None
    lunch = None

    tokens = cell.split('|')
    if len(tokens) < 2:
        raise ValueError('Incorrect number of tokens', tokens)

    tokens = [' '.join(t.split()) for t in tokens]
    start = fix_time_token(tokens[0])
    start = date.combine(date, start)
    end = fix_time_token(tokens[1])
    end = date.combine(date, end)

    if start > end:
        end = end + timedelta(days=1)
    if len(tokens) == 2:
        return [start, end]
    else:
        lunch = fix_time_token(tokens[2])
        lunch = date.combine(date, lunch)
        return [start, end, lunch]


def fix_cell(cell: str) -> str:
    c = None
    tokens = None

    c = cell.strip()
    if c == '*':
        pass
    elif c[:3].lower() in ('off', 'sic', 'vac'):
        c = '*'
    elif not c[0].isnumeric():
        print('Unknown :', c)
        c = '*'
    else:
        c = c.replace('\r\n', ' ').replace('  ', ' ')
        c = c.replace(' LUNCH : ', '|')
        c = c.replace(' - ', '|').replace(' PST', '')
        c = c.replace(' NO LUNCH', '').replace('LUNCH ', '')
        tokens = c.split('|')
        if len(tokens) > 1:
            tokens = [' '.join(t.split()[0:2]) for t in tokens]
            c = '|'.join(tokens)
    return c


def sanitize(htmldoc: BeautifulSoup):
    for table in htmldoc('table'):
        for tag in table.find_all(True):
            tag.attrs.clear()
        for row in table('tr'):
            cells = row('td')
            for c in cells:
                c.string = c.get_text().strip()
                if not c.string:
                    c.string = '*'
            if cells[0].string + cells[1].string == '**':
                row.decompose()
            elif cells[0].string == '*':
                cells[0].string = '@DATES'
            else:
                for c in cells[1:]:
                    c.string = fix_cell(c.string)


def build_schedules(htmldoc: BeautifulSoup, msg_date: datetime) -> list:
    schedules = []
    schedule = None
    dates = None
    for table in htmldoc('table'):
        for row in table('tr'):
            cells = row('td')
            cells = [c.get_text().strip() for c in cells]
            if cells[0] == '@DATES':
                dates = [fix_date_token(d, msg_date) for d in cells[1:]]
                if schedule:
                    schedules.append(schedule)
                schedule = {}
            else:
                name = cells.pop(0)
                for date, cell in zip(dates, cells):
                    if cell == '*':
                        continue
                    try:
                        if not name in schedule:
                            schedule[name] = []
                        schedule[name].append(
                            process_workday(date, cell))
                    except ValueError as e:
                        traceback.print_exc()
    if schedule and len(schedule) > 0:
        schedules.append(schedule)
    return schedules


def main():
    gmail_msgs = None
    msg_date = None
    markup = None
    schedules = []

    with open('messages.pickle', 'rb') as fin:
        gmail_msgs = pickle.load(fin)

    for gmail_msg in gmail_msgs:
        msg_date = extract_date(gmail_msg)
        markup = get_parts('text/html', gmail_msg)
        markup = markup[0]['body']['data']
        markup = base64.urlsafe_b64decode(markup).decode('UTF8')
        htmldoc = BeautifulSoup(markup, features='html.parser')

        sanitize(htmldoc)
        schedules.extend(build_schedules(htmldoc, msg_date))
        # break

    schedules = [s for s in schedules if len(s) > 0]
    # pprint(schedules)
    print('@'*20, len(schedules))
    for sch in schedules:
        print(' '*5, '#'*15, len(sch))
        for emp, wds in sch.items():
            print(' '*10, emp, '-'*10, len(wds))
            # for wd in wds:
            #     print(wd)


if __name__ == "__main__":
    main()

    # service = get_service()
    # # Call the Gmail API
    # gmail_lbls = service.users().labels().list(userId='me').execute()
    # for label in gmail_lbls['labels']:
    #     if label['name'] == 'Horaire':
    #         horaire = label['id']
    #     if label['name'] == 'Processed':
    #         processed = label['id']

    # gmail_msgs = service.users().messages().list(
    #     userId='me', labelIds=[horaire]).execute()

    # for msg in gmail_msgs['messages']:
    #     gmail_msg = service.users().messages().get(
    #         userId='me', id=msg['id'], format='full').execute()
    #     messages.applunch(gmail_msg)

    # with open('messages.pickle', 'wb') as fout:
    #     pickle.dump(messages, fout)

    # exit()
