from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime

import time  # for time sleep
import sys  # for PyCharm compatibility with Windows10 64bit
from apiclient import errors  # for Google API errors
import simplejson  # for Google API errors content

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-rpi-status.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'  # client credentials
APPLICATION_NAME = 'Google Calendar API Python'

sys.modules['win32file'] = None  # needed for PyCharm on Windows10 64bit

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'calendar-python-rpi-status.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z'

    print("\nChecking your calendar for events...")

    eventsResult = service.events().list(
        calendarId='primary', timeMin=now, maxResults=10, singleEvents=True,
        orderBy='startTime').execute()
    events = eventsResult.get('items', [])
    #print(events)

    counter = 0

    for event in events:
        if event['sequence'] is not None:
            counter += 1
    print('Number of running events: ', counter)

    meeting = False
    busy = False

    if not events:
        print('No running events found.')
    counter = 1
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))  # to get date and time of event
        transparency = event.get('transparency')  # to get busy/available status
        summary = event.get('summary')  # to get the title
        description = event.get('description')  # description of the event
        #print('NO DESCRIPTION', description)
        #print(summary.find('meeting'))
        #print(description.find('meeting'))
        #print(start, event['summary'])
        print('Event', counter, end="")
        print(': ', end="")
        if summary is not None:
            print(event['summary'])
        else:
            print('(No title)')
        if summary is None or description is None:
            if transparency not in ['transparent']:
                #print('BUSY')
                busy = True
            #else:
                #print('AVAILABLE')
        elif summary is not None and description is not None:
            if summary.find('meeting') >= 0 or description.find('meeting') >= 0:
                meeting = True
            if transparency not in ['transparent'] and meeting:
                #print('IN A MEETING')
                busy = True
            elif transparency not in ['transparent']:
                #print('BUSY')
                busy = True
            #else:
                #print('AVAILABLE')
        counter += 1
    #print(busy, meeting)
    print('Status of the office: ', end="")
    if busy and meeting:
        print('IN A MEETING')
    elif busy:
        print('BUSY')
    else:
        print('AVAILABLE')

request = 0

if __name__ == '__main__':
    while True:
        try:
            main()
            time.sleep(1)
            request += 1
            print('Requests: ', request)
        except errors.HttpError as err:
            error = simplejson.loads(err.content)
            if error.get('code') == 403:
                print('Rate limit exceeded! Waiting 10 seconds')
                time.sleep(10)
            elif error.get('code') == 500:
                print('Server Internal Error')
                time.sleep(10)
