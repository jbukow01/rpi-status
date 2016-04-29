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
from socket import error as socket_error  # for socket.error
import errno  # for error numbers
from datetime import timedelta  # for threshold time for timeMax

import RPi.GPIO as GPIO  # import of gpios for raspberry pi
GPIO.setwarnings(False)

"""pin numbers setup"""
meeting_pin = 15  # pin on raspberry pi board for meeting light/indicator
busy_pin = 16  # pin for busy light/indicator
available_pin = 18  # pin for available light/indicator
pir_pin = 29  # pin for input from PIR sensor

"""setup modes for pins IN/OUT"""
GPIO.setmode(GPIO.BOARD)
GPIO.setup(meeting_pin, GPIO.OUT)
GPIO.setup(busy_pin, GPIO.OUT)
GPIO.setup(available_pin, GPIO.OUT)
GPIO.setup(pir_pin, GPIO.IN)

"""default start with lights off"""
GPIO.output(meeting_pin, GPIO.LOW)
GPIO.output(busy_pin, GPIO.LOW)
GPIO.output(available_pin, GPIO.LOW)

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# if modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-rpi-status.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'  # client credentials
APPLICATION_NAME = 'Google Calendar API Python'

sys.modules['win32file'] = None  # needed for PyCharm on Windows10 64bit

"""default variables at start"""
status = ''
counter = 0
titles = []
desc_text = []
request = 0
start_time = time.time()
previously_away = None

"""custom variables to modify as required"""
max_events = 5  # maximum number of events the script will check
update_interval = 0  # interval between updates in seconds (increase only if "Rate limit exceeded!" error appears)
error_wait_time = 10  # error wait time interval in seconds before next try
away_time = 0.5  # time in minutes without any movement in the office before the status will change to away
# the below will only display when external screen is attached
meeting_status = 'IN A MEETING'  # text when in a meeting
busy_status = 'BUSY'  # text when busy
available_status = 'AVAILABLE'  # text when available
away_status = 'AWAY'  # test when away

"""calibration of the PIR sensor (no movement should be performed during calibration)"""
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


def get_credentials():
    """gets valid user credentials from local storage.
    if credentials are not present, or if they are invalid,
    the OAuth2 flow is executed to get the new credentials.
    it returns the obtained credential.
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


def switch_lights_off():
    GPIO.output(meeting_pin, GPIO.LOW)
    GPIO.output(busy_pin, GPIO.LOW)
    GPIO.output(available_pin, GPIO.LOW)


def lights_flash():
    GPIO.output(meeting_pin, GPIO.HIGH)
    GPIO.output(busy_pin, GPIO.LOW)
    GPIO.output(available_pin, GPIO.LOW)
    time.sleep(0.1)
    GPIO.output(meeting_pin, GPIO.LOW)
    GPIO.output(busy_pin, GPIO.HIGH)
    GPIO.output(available_pin, GPIO.LOW)
    time.sleep(0.1)
    GPIO.output(meeting_pin, GPIO.LOW)
    GPIO.output(busy_pin, GPIO.LOW)
    GPIO.output(available_pin, GPIO.HIGH)
    time.sleep(0.1)


def detection(start):
    """it checks the movement and returns true
    or false if the away time is met"""
    global start_time
    if not GPIO.input(pir_pin):
        elapsed_time = time.time() - start
        if away_time*20 < elapsed_time < away_time*60:
            switch_lights_off()
            time.sleep(1)
        elif away_time*60 < elapsed_time < away_time*63:
            flash = 0
            while flash < 5:
                lights_flash()
                flash += 1
        elif elapsed_time > away_time*62:
            switch_lights_off()
            return False
        return True
    else:
        start_time = time.time()
        return True


def get_events():
    """it gets the events from the calendar"""
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z'
    threshold = (datetime.datetime.utcnow() + timedelta(seconds=1)).isoformat() + 'Z'

    events_result = service.events().list(
        calendarId='primary', timeMin=now, timeMax=threshold, maxResults=max_events, singleEvents=True,
        orderBy='startTime').execute()
    events = events_result.get('items', [])
    return events


def options():
    """it checks the details about events
    if meeting word is present in the title or description
    it returns true for meeting status
    it returns busy as true only if at least one event is set to "Busy"
    in the calendar
    it also returns obtained new titles, descriptions and running events count"""
    events = get_events()
    meeting = False
    busy = False

    new_counter = 0
    new_titles = []
    new_desc_text = []

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
    return new_titles, new_desc_text, meeting, busy, new_counter


def meeting_on():
    GPIO.output(meeting_pin, GPIO.HIGH)
    GPIO.output(busy_pin, GPIO.LOW)
    GPIO.output(available_pin, GPIO.LOW)


def busy_on():
    GPIO.output(busy_pin, GPIO.HIGH)
    GPIO.output(meeting_pin, GPIO.LOW)
    GPIO.output(available_pin, GPIO.LOW)


def available_on():
    GPIO.output(available_pin, GPIO.HIGH)
    GPIO.output(meeting_pin, GPIO.LOW)
    GPIO.output(busy_pin, GPIO.LOW)


def lights():
    """it checks for motions invoking detection() function
    and it turns the lights off if no motion is present
    or it turns the lights on according to the events"""
    motion_present = detection(start_time)
    new_titles, new_desc_text, meeting, busy, new_counter = options()
    if not motion_present:
        new_status = away_status
        switch_lights_off()
    elif busy and meeting:
        new_status = meeting_status
        meeting_on()
    elif busy:
        new_status = busy_status
        busy_on()
    else:
        new_status = available_status
        available_on()
    return new_status


def status_print():
    """it prints the status of the office in the console"""
    global request
    global status
    global counter
    global titles
    global desc_text
    global previously_away
    new_titles, new_desc_text, meeting, busy, new_counter = options()
    new_status = lights()

    if new_status != status or new_counter != counter or new_titles != titles or new_desc_text != desc_text:
        status = new_status
        counter = new_counter
        titles = new_titles
        desc_text = new_desc_text
        if status == away_status and not previously_away:
            print('\nStatus of the office: ', end="")
            print(status)
            request = 0
            previously_away = True
        elif status != away_status:
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
            previously_away = False

if __name__ == '__main__':
    while True:
        try:
            status_print()
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
        except KeyboardInterrupt:
            print('\nInterrupted by external source')
            break
        except:
            print('Unexpected Error! (UPDATE PENDING...)')
            time.sleep(error_wait_time)

GPIO.cleanup()
