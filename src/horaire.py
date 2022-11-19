"""
	* pyHoraire *
	Prends mon horaire a partir du email et
    va remplir mon calendrier Google
"""

import base64
import os
import pickle
# import traceback
from datetime import date, datetime, time, timedelta
from slugify import slugify
from bs4 import BeautifulSoup
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
    # The file token.pickle stores the user's access and refresh tokens,
    # and is created automatically when the authorization flow completes
    # for the first time.
    creds = None
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
    service = build('gmail', 'v1', credentials=creds)
    return service


def get_parts(mimetype: str, part: object) -> list:
    """ Return the MIME parts of the type sent as parameter

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
        for sub_part in part['parts']:
            parts.extend(get_parts(mimetype, sub_part))
    elif part['mimeType'] == mimetype:
        parts.append(part)
    return parts


def extract_date(gmail_msg: dict) -> date:
    """ Extract the date from the Gmail message

    Args:
        gmail_msg (dict): A gmail message (which is a list)

    Returns:
        date: The date that was extracted
    """
    msg_date = filter(
        lambda h: h['name'] == 'Date', gmail_msg['payload']['headers'])
    msg_date = list(msg_date)[0]['value']
    msg_date = msg_date.split()[1:4]
    msg_date = ' '.join(msg_date)
    msg_date = datetime.strptime(msg_date, "%d %b %Y").date()
    return msg_date


def fix_date_token(dirty: str, msg_date: date) -> datetime:
    """ Fill the missing parts of the datetime

    Args:
        cell (str): The string that contains the incomplete date
        msg_date (date): The date the message was sent

    Returns:
        datetime: Returns the fixed date
    """
    fixed = dirty.split()[1:3]
    fixed.append(str(msg_date.year))
    fixed = ','.join(fixed)
    fixed = datetime.strptime(fixed, '%b,%d,%Y').date()
    if fixed < msg_date:
        # If schedule is for january next year
        fixed = fixed.replace(year=fixed.year + 1)
    return fixed


def fix_time_token(time_token: str) -> time:
    """ Fix and convert the time sent as as string

    Args:
        time_token (str): the string that contains the time

    Returns:
        time: Returns the time fixed
    """
    time_ = None
    if len(time_token) in (4, 5):
        if time_token.find(' ') > -1:
            time_ = datetime.strptime(time_token, '%I %p').time()
        else:
            time_ = datetime.strptime(time_token, '%I%p').time()
    elif len(time_token) in (7, 8):
        time_ = datetime.strptime(time_token, '%I:%M %p').time()
    time_ = time_.replace(tzinfo=PST)
    return time_


def process_workday(cell: str, date_: date) -> dict:
    """ Get the tokens from the string and add the time to the date

    Args:
        date (datetime): The date part of the workhours
        cell (str): the string that contains the workhours

    Raises:
        ValueError: if the cell is not well formatted and
        not enough tokens are found

    Returns:
        dict: Returns a dictionary with the workhours
    """
    tokens = cell.split('|')
    if not len(tokens) in (2, 3):
        raise ValueError('Incorrect number of tokens', tokens)

    tokens = [' '.join(t.split()) for t in tokens]
    start = fix_time_token(tokens[0])
    start = datetime.combine(date_, start)
    end = fix_time_token(tokens[1])
    end = datetime.combine(date_, end)

    if start > end:
        end = end + timedelta(days=1)
    if len(tokens) == 2:
        workhours = {'start': start, 'end': end}
    else:
        lunch = fix_time_token(tokens[2])
        lunch = datetime.combine(date_, lunch)
        workhours = {'start': start, 'end': end, 'lunch': lunch}
    return workhours


def fix_cell(dirty: str) -> str:
    """ Clean up the string

    Args:
        dirty (str): The string to clean up

    Returns:
        str: The string cleaned up
    """
    tokens = None
    fixed = dirty.strip()
    if fixed == '*':
        pass
    elif fixed[:3].lower() in ('off', 'sic', 'vac'):
        fixed = '*'
    elif not fixed[0].isnumeric():
        print('Unknown :', fixed)
        fixed = '*'
    else:
        fixed = fixed.replace('\r\n', ' ').replace('  ', ' ')
        fixed = fixed.replace(' LUNCH : ', '|')
        fixed = fixed.replace(' - ', '|').replace(' PST', '')
        fixed = fixed.replace(' NO LUNCH', '').replace('LUNCH ', '')
        tokens = fixed.split('|')
        if len(tokens) > 1:
            tokens = [' '.join(t.split()[0:2]) for t in tokens]
            fixed = '|'.join(tokens)
    return fixed


def sanitize(htmldoc: BeautifulSoup):
    """ Sanitize the html to make it easier to parse

    Args:
        htmldoc (BeautifulSoup): The html document to sanitize
    """
    for table in htmldoc('table'):
        for tag in table.find_all(True):
            tag.attrs.clear()
        for row in table('tr'):
            cells = row('td')
            for cell in cells:
                cell.string = cell.get_text().strip()
                if not cell.string:
                    cell.string = '*'
            if cells[0].string + cells[1].string == '**':
                row.decompose()
            elif cells[0].string == '*':
                cells[0].string = '@DATES'
            else:
                for cell in cells[1:]:
                    cell.string = fix_cell(cell.string)


def build_schedules(htmldoc: BeautifulSoup, msg_date: datetime) -> list:
    """ Take the htmldoc, parse it and return the schedules

    Args:
        htmldoc (BeautifulSoup): the html document to be parsed
        msg_date (datetime): The date the message was received to help
            determine the dates in the schedule

    Returns:
        list: Returns the parsed results
    """
    schedules = []
    schedule = {}
    dates = None
    for row in htmldoc('tr'):
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
            for cell, date_ in zip(cells, dates):
                if cell == '*':
                    continue
                if not name in schedule:
                    schedule[name] = []
                try:
                    schedule[name].append(
                        process_workday(cell, date_))
                except ValueError:
                    # traceback.print_exc()
                    pass
    return schedules


def create_event(dt_start: datetime,
                 dt_end: datetime, uid: str, summary) -> Event:
    """ Return a new event

    Args:
        dt_start (datetime): The date at which the event will start

        dt_end (datetime):  The date at which the event will end
                            uid (str): The ID of the event

        summary ([type]): Description for the event

    Returns:
        Event: [description]
    """
    event_ = Event()
    event_.add('UID', uid)
    event_.add('DTSTART', dt_start)
    event_.add('DTEND', dt_end)
    event_.add('SUMMARY', summary)
    event_.add('DTSTAMP', datetime.now())
    return event_


def create_alarm(diff: timedelta, related: str) -> Alarm:
    """ Return a new alarm

    Args:
        diff (timedelta): The delta between the alarm and
                        the start/end of the event
        related (str):  Is the delta related to the START
                        or the END of the event

    Returns:
        Alarm: The Alarm itself
    """
    alarm = Alarm()
    alarm.add('ACTION', 'DISPLAY')
    alarm.add('TRIGGER', diff, parameters={'RELATED': related})
    return alarm


def turn_to_ical(slug: str, workdays: list) -> bytes:
    """ Take the schedule and make a iCalendar out of it

    Args:
        slug (str): The slug to use in the UID of the events/alarms
                    Probably the name of the employee
        workdays (list): The workdays to transform into events

    Returns:
        bytes: The calendar in iCal format
    """
    ical = Calendar()
    ical.add('prodid', '-//pyhoraire//')
    ical.add('version', '2.0')
    for workday in workdays:
        dt_start = workday['start']
        dt_end = workday['end']
        uid = dt_start.timestamp()
        uid = f'w/{uid}/{slug}/geeksquad.ca'
        summary = 'Travail'
        ev_workday = create_event(dt_start, dt_end, uid, summary)
        alarm = create_alarm(timedelta(minutes=-15), 'START')
        ev_workday.add_component(alarm)
        alarm = create_alarm(timedelta(minutes=-5), 'START')
        ical.add_component(ev_workday)
        if 'lunch' in workday:
            dt_start = workday['lunch']
            dt_end = dt_start + timedelta(minutes=30)
            uid = dt_start.timestamp()
            uid = f'l/{uid}/{slug}/geeksquad.ca'
            summary = 'Lunch'
            ev_lunch = create_event(dt_start, dt_end, uid, summary)
            alarm = create_alarm(timedelta(minutes=-5), 'START')
            ev_lunch.add_component(alarm)
            alarm = create_alarm(timedelta(minutes=-5), 'END')
            ev_lunch.add_component(alarm)
            ical.add_component(ev_lunch)
    return ical.to_ical(sorted=True)


def output_schedules(schedules: list) -> None:
    """ Output the schedules to the screen

    Args:
        schedules (list): The schedules to be output
    """
    print('@'*20, len(schedules))
    for schedule in schedules:
        print(' '*5, '#'*15, len(schedule))
        for employee, workdays in schedule.items():
            print(' '*10, employee, '-'*10, len(workdays))
            for workday in workdays:
                print(workday)


def export_schedules(schedules: list) -> None:
    """ Create the .ics files from the schedules

    Args:
        schedules (list): The schedules to be exported
    """
    for sch in schedules:
        for employee, workdays in sch.items():
            if employee == '@DATES':
                continue
            if not employee.startswith('Ste-Marie'):
                continue
            employee_slug = slugify(employee)
            ical = turn_to_ical(employee_slug, workdays)
            dt_start = sch['@DATES'][0].strftime('%Y-%m-%d')
            dt_end = sch['@DATES'][6].strftime('%Y-%m-%d')
            filename = f'Schedule {dt_start} to {dt_end}.ics'
            filename = os.path.join('calendars',
                                    employee_slug, filename)
            os.makedirs(os.path.join('calendars',
                                     employee_slug), exist_ok=True)
            with open(filename, 'wb') as fout:
                fout.write(ical)


def main():
    """ Get emailed schedules, parse them and build an icalendar"""
    gmail_msgs = []
    schedules = []
    processed_id = None
    unprocessed_id = None
    service = get_service()

    # Call the Gmail API to get the user's labels
    results = service.users().labels().list(userId='me').execute() \
        # pylint: disable=no-member
    # Identify the labels used for the schedules
    for label in results['labels']:
        if label['name'] == UNPROCESSED_LABEL:
            unprocessed_id = label['id']
        if label['name'] == PROCESSED_LABEL:
            processed_id = label['id']

    # Get the IDs of messages that have the label applied
    results = service.users().messages().list(  # pylint: disable=no-member
        userId='me', labelIds=[unprocessed_id]).execute()

    # Get the body of those messages
    for msg in results['messages']:
        msg = service.users().messages().get(  # pylint: disable=no-member
            userId='me', id=msg['id'], format='full').execute()

        gmail_msgs.append(msg)
        gmail_msgs = [m for m in gmail_msgs if
                      not processed_id in m['labelIds']]

    # with open('messages.pickle', 'wb') as fout:
    #     pickle.dump(gmail_msgs, fout)
    # exit()
    # with open('messages.pickle', 'rb') as fin:
    #     gmail_msgs = pickle.load(fin)

    # For each message, get the html inside
    for gmail_msg in gmail_msgs:
        msg_date = extract_date(gmail_msg)
        markup = get_parts('text/html', gmail_msg)[0]['body']['data']
        # markup = markup[0]['body']['data']
        markup = base64.urlsafe_b64decode(markup).decode('UTF8')
        htmldoc = BeautifulSoup(markup, features='html.parser')

        # Prepare the html to be processed
        sanitize(htmldoc)
        # Extract the schedules from the html
        schedules.extend(build_schedules(htmldoc, msg_date))

    # output_schedules(schedules)
    export_schedules(schedules)


if __name__ == "__main__":
    main()
