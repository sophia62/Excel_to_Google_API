from excel_reader import process_excel
from google_calendar_integration import get_calendar_service, create_event

def main():
    # Update the filename to your actual Excel filename
    events = process_excel('Study_Fall_2024.xlsx')
    service = get_calendar_service()
    for event_details in events:
        create_event(service, event_details)

if __name__ == "__main__":
    main()
