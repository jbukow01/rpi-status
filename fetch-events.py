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
import ssl  # only for ssl.SSLEOError handling
#import errno
#from socket import error as error_socket
#from time import gmtime, strftime
from datetime import timedelta

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

status = ''
counter = 0
titles = []
description_text = []

""" custom variables """
maxEvents = 5  # maximum number of events the script will check
updateInterval = 1  # interval between updates in seconds
errorWaitTime = 10  # error wait time interval in seconds
meetingStatus = 'IN A MEETING'  # text when in a meeting
busyStatus = 'BUSY'  # text when busy
availableStatus = 'AVAILABLE'  # text when available


def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z'
    #current_time = strftime("%Y-%m-%dT%H:%M:%S", gmtime())
    #current_time = (datetime.datetime.now()).isoformat()
    threshold = (datetime.datetime.utcnow() + timedelta(seconds=1)).isoformat() + 'Z'

    eventsResult = service.events().list(
        calendarId='primary', timeMin=now, timeMax=threshold, maxResults=maxEvents, singleEvents=True,
        orderBy='startTime').execute()
    events = eventsResult.get('items', [])
    #print(events)

    #for event in events:
    #    if event['sequence'] is not None:
    #        counter += 1

    meeting = False
    busy = False

    #if not events:
        # print('No running events found.')

    new_counter = 0
    new_titles = []
    new_description_text = []
    #titles = []
    #description_text = []
    #message = {}
    message_text = {}
    message_text['checking'] = '\nChecking your calendar for events...'

    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))  # to get start date and time of event
        end = event['end'].get('dateTime', event['end'].get('date'))  # to get end date and time of event
        transparency = event.get('transparency')  # to get busy/available status
        summary = event.get('summary')  # to get the title
        description = event.get('description')  # description of the event
        # print('NO DESCRIPTION', description)
        # print(summary.find('meeting'))
        # print(description.find('meeting'))
        # print(start, event['summary'])
        # print('Event', counter, end="")
        # print(': ', end="")
        if summary is not None:
            new_titles.append(summary)
            #message['title'] = summary
            lower_case = summary.lower()
            if lower_case.find('meeting') >= 0:
                meeting = True
        else:
            new_titles.append('(No title)')
            #message['title'] = '(No title)'
        if description is not None:
            new_description_text.append(description)
            #print(description)
            lower_case = description.lower()
            if lower_case.find('meeting') >= 0:
                meeting = True
        else:
            new_description_text.append('(No description)')
        if transparency not in ['transparent']:
            busy = True
        new_counter += 1
    # print(busy, meeting)
    """
    print("\nChecking your calendar for events...")
    print('Number of running events: ', counter)
    no_titles = 0
    for title in titles:
        no_titles += 1
        print('Event', no_titles, end="")
        print(': ', end="")
        print(title)
    print('Status of the office: ', end="")
    """
    if busy and meeting:
        new_status = meetingStatus
        #print('IN A MEETING')
    elif busy:
        new_status = busyStatus
        #print('BUSY')
    else:
        new_status = availableStatus
        #print('AVAILABLE')

    global status
    global counter
    global titles
    global description_text

    if new_status != status or new_counter != counter or new_titles != titles or new_description_text != description_text:
        status = new_status
        counter = new_counter
        new_counter = 0
        titles = new_titles
        description_text = new_description_text
        print("\nChecking your calendar for events...")
        print('Number of running events: ', counter)
        no_titles = 0
        for title in titles:
            no_titles += 1
            print('Event', no_titles, end="")
            print(': ', end="")
            print(title)
            print('Description', no_titles, end="")
            print(': ', end="")
            print(description_text[no_titles - 1])
        print('Status of the office: ', end="")
        print(status)
        print('Last updated: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        print('Number of requests since last update: ', request)
    """
    message_text['checking'] = '\nChecking your calendar for events...'
    message_text['number'] = 'Number of running events: '
    message_text['counter'] = counter
    no_titles = 0
    for title in titles:
        no_titles += 1
        message_text['no_titles'] = no_titles
        message_text['event'[no_titles]] = title
    message_text['status'] = status
    print(message_text)
    """

request = 0

if __name__ == '__main__':
    while True:
        try:
            main()
            time.sleep(updateInterval)
            request += 1
            #print('Requests: ', request)
        except errors.HttpError as err:
            error = simplejson.loads(err.content)
            if error.get('code') == 403 and \
                error.get('errors')[0].get('reason') \
                    in ['rateLimitExceeded', 'userRateLimitExceeded']:
                print('Rate limit exceeded! Waiting 10 seconds...')
                time.sleep(errorWaitTime)
            elif error.get('code') == 500:
                print('Server Internal Error (UPDATE PENDING...)')
                time.sleep(errorWaitTime)
            elif IOError:
                print('I/O error (UPDATE PENDING...)')
                time.sleep(errorWaitTime)
            elif ssl.SSLError:
                print('SSL error (UPDATE PENDING...)')
                time.sleep(errorWaitTime)
            else:
                print('Unexpected Error! (UPDATE PENDING...)')
                time.sleep(errorWaitTime)
