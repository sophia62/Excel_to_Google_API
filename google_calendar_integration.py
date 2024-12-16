from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime

SCOPES = ['https://www.googleapis.com/auth/calendar']
SERVICE_ACCOUNT_FILE = 'credentials.json'

def get_calendar_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('calendar', 'v3', credentials=creds)
    return service

def create_event(service, event_details):
    try:
        start_date = datetime.date.fromisoformat(event_details['Date'])
        start_time = datetime.time.fromisoformat(event_details['Time'])
    except ValueError:
        return

    start_datetime = datetime.datetime.combine(start_date, start_time)
    end_datetime = start_datetime + datetime.timedelta(hours=1)

    event = {
        'summary': event_details['Description'],
        'location': event_details.get('Location', ''),
        'start': {
            'dateTime': start_datetime.isoformat(),
            'timeZone': 'America/New_York',
        },
        'end': {
            'dateTime': end_datetime.isoformat(),
            'timeZone': 'America/New_York',
        }
    }

    event_result = service.events().insert(calendarId='sophiabeebe1@gmail.com', body=event).execute()

    print(f"Event created: {event_result.get('htmlLink')}")
