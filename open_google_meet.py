from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import webbrowser
import time
import pytz
import pyttsx3


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
SLEEP_WINDOW_SECS = 60 * 20
OPEN_MEETING_MINUTES_BEFORE = 3
speech_engine = pyttsx3.init()


def initiate_credentials():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def get_next_meeting_details():
    creds = initiate_credentials()
    service = build('calendar', 'v3', credentials=creds)
    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    now_plus_window = (datetime.datetime.utcnow() + datetime.timedelta(hours=8)).isoformat() + 'Z'
    events_result = service.events().list(calendarId='primary', timeMin=now, timeMax=now_plus_window,
                                        maxResults=10, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start_time = datetime.datetime.fromisoformat(event['start'].get('dateTime', event['start'].get('date')))
        print(start_time, event['summary'])
        meeting_link = event.get('hangoutLink')
        if meeting_link and is_time_in_future(start_time):
            return {'meeting_link': meeting_link,
                    'start_time': start_time,
                    'meeting_name': event.get('summary')}
    return {}

def is_time_in_future(time_to_check):
    return time_to_check > datetime.datetime.utcnow().replace(tzinfo=pytz.utc)

def open_meeting_in_browser(meeting_link):
    print(f"Initiating meeting {meeting_link}")
    webbrowser.open(meeting_link)

def get_secs_till_next_meeting(meeting_start_time):
    if meeting_start_time:
        return (meeting_start_time - datetime.datetime.utcnow().replace(tzinfo=pytz.utc) -
               datetime.timedelta(minutes=OPEN_MEETING_MINUTES_BEFORE)).seconds
    return SLEEP_WINDOW_SECS + 10

def alert_on_meeting(meeting_name):
    alert_message = f"{meeting_name} will start in {OPEN_MEETING_MINUTES_BEFORE} minutes"
    speech_engine.say(alert_message)
    speech_engine.runAndWait()

def main():
    meeting_should_be_shown_soon = False
    meeting_link = None
    secs_till_next_meeting = SLEEP_WINDOW_SECS + 10
    meeting_name = ''

    while True:
        # showing the next meeting
        if meeting_should_be_shown_soon and meeting_link:
            open_meeting_in_browser(meeting_link)
            meeting_should_be_shown_soon = False
            meeting_link = None
            alert_on_meeting(meeting_name)
            meeting_name = ''
        # getting the next meeting
        else:
            meeting_details = get_next_meeting_details()
            meeting_link = meeting_details.get('meeting_link')
            meeting_start_time = meeting_details.get('start_time')
            meeting_name = meeting_details.get('meeting_name')
            secs_till_next_meeting = get_secs_till_next_meeting(meeting_start_time)
        if SLEEP_WINDOW_SECS > secs_till_next_meeting:
            meeting_should_be_shown_soon = True
        time.sleep(min(SLEEP_WINDOW_SECS, secs_till_next_meeting))

if __name__ == '__main__':
    main()