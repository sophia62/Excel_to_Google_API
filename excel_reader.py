import pandas as pd
from dateutil import parser
import datetime

def get_next_weekday(weekday_name):
    weekday_name = weekday_name.strip().upper()
    weekdays = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
    if weekday_name not in weekdays:
        return None
    today = datetime.date.today()
    target_index = weekdays.index(weekday_name)
    days_ahead = (target_index - today.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7  # schedule for next week if it's the same day
    return today + datetime.timedelta(days=days_ahead)

def process_excel(filename):
    try:
        df = pd.read_excel(filename)
    except FileNotFoundError:
        print(f"Error: The file '{filename}' was not found.")
        return []
    except Exception as e:
        print(f"Error reading the Excel file '{filename}': {e}")
        return []

    # Convert columns to upper case for consistency
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Check for TIME column
    if "TIME" not in df.columns:
        print("Warning: No 'TIME' column found in the Excel file. No events will be scheduled.")
        return []

    # Identify day columns
    day_columns = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
    # Filter to columns that exist in the file
    existing_days = [d for d in day_columns if d in df.columns]

    if not existing_days:
        print("Warning: No weekday columns (MONDAY-SUNDAY) found in the Excel file. No events will be scheduled.")
        return []

    events = []
    for index, row in df.iterrows():
        time_str = row.get("TIME")
        if pd.isna(time_str):
            # If no time, skip this row
            continue

        # Try to parse the time
        try:
            parsed_time = parser.parse(str(time_str))
            event_time = parsed_time.strftime("%H:%M:%S")
        except Exception:
            # If time parsing fails, skip this row
            print(f"Warning: Unable to parse time '{time_str}' in row {index}. Skipping this row.")
            continue

        # Iterate over each day column
        for day in existing_days:
            event_desc = row.get(day)
            if pd.notna(event_desc) and str(event_desc).strip():
                event_date = get_next_weekday(day)
                if event_date is None:
                    continue
                events.append({
                    'Description': str(event_desc).strip(),
                    'Date': event_date.isoformat(),
                    'Time': event_time,
                    'Location': "",
                    'Priority': ""
                })

    if not events:
        print("No events found to schedule.")
    return events
