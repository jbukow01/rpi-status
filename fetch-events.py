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
import socket  # for socket.error
import errno
from socket import error as socket_error
from datetime import timedelta

import RPi.GPIO as GPIO # import of gpios for raspberry pi
GPIO.setwarnings(False)

meeting_pin = 15
busy_pin = 16
available_pin = 18
pir_pin = 29

GPIO.setmode(GPIO.BOARD)
GPIO.setup(meeting_pin, GPIO.OUT)
GPIO.setup(busy_pin, GPIO.OUT)
GPIO.setup(available_pin, GPIO.OUT)
GPIO.setup(pir_pin, GPIO.IN)

GPIO.output(meeting_pin, GPIO.LOW)
GPIO.output(busy_pin, GPIO.LOW)
GPIO.output(available_pin, GPIO.LOW)

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
request = 0
desc_text = []

""" custom variables """
max_events = 5  # maximum number of events the script will check
update_interval = 1  # interval between updates in seconds
error_wait_time = 10  # error wait time interval in seconds
meeting_status = 'IN A MEETING'  # text when in a meeting
busy_status = 'BUSY'  # text when busy
available_status = 'AVAILABLE'  # text when available
away_status = 'AWAY'
away_time = 0.2  # time in minutes without any movement in the office before the status will change to away

calibration = 0
print("Motion sensor calibration", end='')
sys.stdout.flush()
while calibration < 20:
    print(".", end='')
    sys.stdout.flush()
    GPIO.output(meeting_pin, GPIO.HIGH)
    calibration += 1
    time.sleep(0.2)
    GPIO.output(meeting_pin, GPIO.LOW)
    time.sleep(0.2)
print("\nStatus in operation!")


def detection(start_time):
    while not GPIO.input(pir_pin):
        elapsed_time = time.time() - start_time
        if elapsed_time > away_time*60:
            return False
    return True


def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z'
    threshold = (datetime.datetime.utcnow() + timedelta(seconds=1)).isoformat() + 'Z'

    events_result = service.events().list(
        calendarId='primary', timeMin=now, timeMax=threshold, maxResults=max_events, singleEvents=True,
        orderBy='startTime').execute()
    events = events_result.get('items', [])

    meeting = False
    busy = False

    global status
    global counter
    global titles
    global desc_text
    global request

    new_counter = 0
    new_titles = []
    new_desc_text = []
    message_text = {}
    message_text['checking'] = '\nChecking your calendar for events...'

    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))  # to get start date and time of event
        end = event['end'].get('dateTime', event['end'].get('date'))  # to get end date and time of event
        transparency = event.get('transparency')  # to get busy/available status
        summary = event.get('summary')  # to get the title
        description = event.get('description')  # description of the event
        if summary is not None:
            new_titles.append(summary)
            lower_case = summary.lower()
            if lower_case.find('meeting') >= 0:
                meeting = True
        else:
            new_titles.append('(No title)')
        if description is not None:
            new_desc_text.append(description)
            lower_case = description.lower()
            if lower_case.find('meeting') >= 0:
                meeting = True
        else:
            new_desc_text.append('(No description)')
        if transparency not in ['transparent']:
            busy = True
        new_counter += 1

    motion_present = detection(time.time())

    if not motion_present:
        new_status = away_status
        GPIO.output(meeting_pin, GPIO.LOW)
        GPIO.output(busy_pin, GPIO.LOW)
        GPIO.output(available_pin, GPIO.LOW)
    elif busy and meeting:
        new_status = meeting_status
        GPIO.output(meeting_pin, GPIO.HIGH)
        GPIO.output(busy_pin, GPIO.LOW)
        GPIO.output(available_pin, GPIO.LOW)
    elif busy:
        new_status = busy_status
        GPIO.output(busy_pin, GPIO.HIGH)
        GPIO.output(meeting_pin, GPIO.LOW)
        GPIO.output(available_pin, GPIO.LOW)
    else:
        new_status = available_status
        GPIO.output(available_pin, GPIO.HIGH)
        GPIO.output(meeting_pin, GPIO.LOW)
        GPIO.output(busy_pin, GPIO.LOW)

    if new_status != status or new_counter != counter or new_titles != titles or new_desc_text != desc_text:
        status = new_status
        counter = new_counter
        new_counter = 0
        titles = new_titles
        desc_text = new_desc_text
        if status == away_status:
            print('\nStatus of the office: ', end="")
            print(status)
        else:
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
                print(desc_text[no_titles - 1])
            print('Status of the office: ', end="")
            print(status)
            print('Last updated: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            print('Number of requests since previous update: ', request)
            request = 0

if __name__ == '__main__':
    while True:
        try:
            main()
            time.sleep(update_interval)
            request += 1
        except errors.HttpError as err:
            error = simplejson.loads(err.content)
            if error.get('code') == 403 and \
                error.get('errors')[0].get('reason') \
                    in ['rateLimitExceeded', 'userRateLimitExceeded']:
                print('Rate limit exceeded! Waiting 10 seconds...')
                time.sleep(error_wait_time)
            elif error.get('code') == 500:
                print('Server Internal Error (UPDATE PENDING...)')
                time.sleep(error_wait_time)
            elif IOError:
                print('I/O error (UPDATE PENDING...)')
                time.sleep(error_wait_time)
            elif ssl.SSLError:
                print('SSL error (UPDATE PENDING...)')
                time.sleep(error_wait_time)
            elif ssl.SSLEOFError:
                print('SSL error (UPDATE PENDING...)')
                time.sleep(error_wait_time)
            elif socket.error:
                print('Network is unreachable (UPDATE PENDING...)')
                time.sleep(error_wait_time)
            else:
                print('Unexpected Error! (UPDATE PENDING...)')
                time.sleep(error_wait_time)
        except socket_error as serr:
            if serr.errno == errno.ECONNREFUSED:
                print('Connection Refused! (UPDATE PENDING...)')
                time.sleep(error_wait_time)
            else:
                print('Unexpected Error! (UPDATE PENDING...)')
                time.sleep(error_wait_time)
        except httplib2.ServerNotFoundError:
            print('Server Not Found! Please check internet connection.')
            time.sleep(error_wait_time)
        except:
            print('Unexpected Error! (UPDATE PENDING...)')
            time.sleep(error_wait_time)

GPIO.cleanup()
