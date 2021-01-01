"""
	* pyHoraire *
	Prends mon horaire a partir du email et va remplir mon calendrier Google 
"""

# If modifying these scopes, delete the file token.pickle.


import base64
import os
import pickle
# import traceback
# from pprint import pprint
from datetime import date, datetime, time, timedelta
from slugify import slugify
from bs4 import BeautifulSoup
from bs4.element import NavigableString, Tag
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from icalendar import Calendar, Event, Alarm
from pytz import timezone

SCOPES = ['https://mail.google.com/']
PST = timezone('Canada/Pacific')
UNPROCESSED_LABEL = 'Horaire'
PROCESSED_LABEL = 'Processed'


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
    return build('gmail', 'v1', credentials=creds)


def get_parts(mimetype: str, part: object) -> list:
    """ Goes through the MIME part sent and returns the MIME
        parts that are of the type sent as parameter. Used recursively

    Args:
        mimetype (str): the MIME types we're looking for
        part (object): The part of the email to look in

    Returns:
        list: A list of the parts found
    """
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


def extract_date(gmail_msg: dict) -> date:
    """Extract the date the message was sent from the message object

    Args:
        gmail_msg (dict): A gmail message (which is a list)

    Returns:
        date: The date that was extracted
    """
    msg_date = filter(
        lambda h: h['name'] == 'Date',
        gmail_msg['payload']['headers'])
    msg_date = list(msg_date)[0]['value']
    msg_date = msg_date.split()[1:4]
    msg_date = ' '.join(msg_date)
    msg_date = datetime.strptime(msg_date, "%d %b %Y").date()
    return msg_date


def fix_date_token(cell: str, msg_date: date) -> datetime:
    """Take the date for which parts are missing.
    Fills the missing parts

    Args:
        cell (str): The string that contains the incomplete date
        msg_date (date): The date the message was sent

    Returns:
        datetime: Returns the fixed date
    """
    dt = cell.split()[1:3]
    dt.append(str(msg_date.year))
    dt = ','.join(dt)
    dt = datetime.strptime(dt, '%b,%d,%Y').date()
    if dt < msg_date:
        # If schedule is for january next year
        dt = dt.replace(year=dt.year + 1)
    return dt


def fix_time_token(time_token: str) -> time:
    """Fixes the time sent as parameter. Tries to recover from mistakes

    Args:
        time_token (str): the string that contains the time

    Returns:
        time: Returns the time fixed
    """
    tm = None
    if len(time_token) in (4, 5):
        if time_token.find(' ') > -1:
            tm = datetime.strptime(time_token, '%I %p').time()
        else:
            tm = datetime.strptime(time_token, '%I%p').time()
    elif len(time_token) in (7, 8):
        tm = datetime.strptime(time_token, '%I:%M %p').time()
    tm = tm.replace(tzinfo=PST)
    return tm


def process_workday(dt: date, cell: str) -> dict:
    """Get the tokens from the cell and add the time to the date

    Args:
        date (datetime): The date part of the workhours
        cell (str): the string that contains the workhours

    Raises:
        ValueError: if the cell is not well formatted and
        not enough tokens are found

    Returns:
        dict: Returns a dictionary with the workhours
    """
    start = None
    end = None
    lunch = None
    tokens = cell.split('|')
    if len(tokens) < 2:
        raise ValueError('Incorrect number of tokens', tokens)

    tokens = [' '.join(t.split()) for t in tokens]
    start = fix_time_token(tokens[0])
    start = datetime.combine(dt, start)
    end = fix_time_token(tokens[1])
    end = datetime.combine(dt, end)

    if start > end:
        end = end + timedelta(days=1)
    if len(tokens) == 2:
        return {'start': start, 'end': end}
    else:
        lunch = fix_time_token(tokens[2])
        lunch = datetime.combine(dt, lunch)
        return {'start': start, 'end': end, 'lunch': lunch}


def fix_cell(cell: str) -> str:
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
    """Sanitizes the html to make it easier to parse

    Args:
        htmldoc (BeautifulSoup): The html document to sanitize
    """
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
    """Takes the html document and parses it return the
        information in a list

    Args:
        htmldoc (BeautifulSoup): the html document to be parsed
        msg_date (datetime): The date the message was received to help
            determine the dates in the schedule

    Returns:
        list: Returns the parsed results
    """
    schedules = []
    schedule = []
    dates = None
    for table in htmldoc('table'):
        for row in table('tr'):
            cells = row('td')
            cells = [c.get_text().strip() for c in cells]
            if cells[0] == '@DATES':
                # The row contains the dates
                dates = [fix_date_token(d, msg_date) for d in cells[1:]]
                if schedule:
                    schedules.append(schedule)
                schedule = {'@DATES': dates}
            else:
                # Get the name of the employee
                name = cells.pop(0)
                for date, cell in zip(dates, cells):
                    if cell == '*':
                        continue
                    try:
                        if not name in schedule:
                            schedule[name] = []
                        schedule[name].append(
                            process_workday(date, cell))
                    except ValueError:
                        # traceback.print_exc()
                        pass
    if schedule and len(schedule) > 0:
        schedules.append(schedule)
    return schedules


def main():
    """Gets the schedules sent to my Gmail, parses them
        and build my calendar
    """
    gmail_msg = None
    gmail_msgs = []
    htmldoc = None
    msg_date = None
    # processed_id = None
    # results = None
    schedules = []
    # unprocessed_id = None

    # service = get_service()
    # # Call the Gmail API to get the user's labels
    # results = service.users().labels().list(userId='me').execute()
    # # Identify the labels used for the schedules
    # for label in results['labels']:
    #     if label['name'] == UNPROCESSED_LABEL:
    #         unprocessed_id = label['id']
    #     if label['name'] == PROCESSED_LABEL:
    #         processed_id = label['id']

    # # Get the IDs of messages that have the label applied
    # results = service.users().messages().list(
    #     userId='me', labelIds=[unprocessed_id]).execute()

    # # Get the body of those messages
    # for msg in results['messages']:
    #     msg = service.users().messages().get(
    #         userId='me', id=msg['id'], format='full').execute()
    #     gmail_msgs.append(msg)
    # # gmail_msgs = [m for m in gmail_msgs if
    # #     not processed_id in m['labelIds']]

    # with open('messages.pickle', 'wb') as fout:
    #     pickle.dump(gmail_msgs, fout)
    # exit()

    with open('messages.pickle', 'rb') as fin:
        gmail_msgs = pickle.load(fin)

    # For each message, get the html inside
    for gmail_msg in gmail_msgs:
        msg_date = extract_date(gmail_msg)
        markup = get_parts('text/html', gmail_msg)
        markup = markup[0]['body']['data']
        markup = base64.urlsafe_b64decode(markup).decode('UTF8')
        htmldoc = BeautifulSoup(markup, features='html.parser')

        # Prepare the html to be processed
        sanitize(htmldoc)
        # Extract the schedules from the html
        schedules.extend(build_schedules(htmldoc, msg_date))

    # print('@'*20, len(schedules))
    # for sch in schedules:
    #     print(' '*5, '#'*15, len(sch))
    #     for emp, wds in sch.items():
    #         print(' '*10, emp, '-'*10, len(wds))
    #         for wd in wds:
    #             print(wd)

    # for sch in schedules[-3:-1]:
    for sch in schedules:
        for employee, workdays in sch.items():
            cal = Calendar()
            cal.add('prodid', '-//pyhoraire//')
            cal.add('version', '2.0')
            if employee == '@DATES':
                continue
            # if not employee.startswith('Ste-Marie'):
            #     continue
            employee_slug = slugify(employee)
            for workday in workdays:
                uid = workday['start'].timestamp()
                uid = f'w/{uid}/{employee_slug}/geeksquad.ca'
                ev_workday = Event()
                ev_workday.add('UID', uid)
                ev_workday.add('DTSTART', workday['start'])
                ev_workday.add('DTEND', workday['end'])
                alarm = Alarm()
                alarm.add('ACTION', 'DISPLAY')
                alarm.add('TRIGGER', timedelta(minutes=-15),
                          parameters={'RELATED': 'START'})
                ev_workday.add_component(alarm)
                alarm = Alarm()
                alarm.add('ACTION', 'DISPLAY')
                alarm.add('TRIGGER', timedelta(minutes=-5),
                          parameters={'RELATED': 'START'})
                ev_workday.add_component(alarm)
                cal.add_component(ev_workday)
                if 'lunch' in workday:
                    uid = workday['lunch'].timestamp()
                    uid = f'l/{uid}/{employee_slug}/geeksquad.ca'
                    ev_lunch = Event()
                    ev_lunch.add('UID', uid)
                    ev_lunch.add('DTSTART', workday['lunch'])
                    ev_lunch.add('DTEND', workday['lunch'] +
                                 timedelta(minutes=30))
                    alarm = Alarm()
                    alarm.add('ACTION', 'DISPLAY')
                    alarm.add('TRIGGER', timedelta(minutes=-5),
                              parameters={'RELATED': 'START'})
                    ev_lunch.add_component(alarm)
                    alarm = Alarm()
                    alarm.add('ACTION', 'DISPLAY')
                    alarm.add('TRIGGER', timedelta(minutes=-5),
                              parameters={'RELATED': 'END'})
                    ev_lunch.add_component(alarm)
                    cal.add_component(ev_lunch)

            dtstart = sch['@DATES'][0].strftime('%Y-%m-%d')
            dtend = sch['@DATES'][6].strftime('%Y-%m-%d')
            filename = f'Schedule {dtstart} to {dtend}.ics'
            filename = os.path.join('calendars',
                                    employee_slug, filename)
            os.makedirs(os.path.join('calendars',
                                        employee_slug), exist_ok=True)
            with open(filename, 'wb') as fout:
                fout.write(cal.to_ical())


if __name__ == "__main__":
    main()
