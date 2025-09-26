import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta
from io import StringIO

def format_activity_name(activity):
    return ''.join(word.capitalize() for word in activity.split('_'))

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

Athlete_TLA = "RAA"
ATP_year = "2026"
ATP_sheet_name = "ATP_Data"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\ATP2intervals_{Athlete_TLA}_{ATP_year}.xlsm"
RACE_file_path = rf"C:\TEMP\{Athlete_TLA}\race_events.xlsx"

parse_delay = .01

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")

url_base = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
url_activities = f"https://intervals.icu/api/v1/athlete/{athlete_id}/activities"
API_headers = {"Content-Type": "application/json"}

oldest_date = f"{ATP_year}-01-01T00:00:00"
newest_date = f"{ATP_year}-12-31T23:59:59"

def get_race_events(athlete_id, username, api_key, race_categories, oldest_date, newest_date):
    url_get = f"{url_base}/events"
    events_list = []
    for category in race_categories:
        params = {
            "oldest": oldest_date,
            "newest": newest_date,
            "category": category
        }
        response = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
        if response.status_code == 200:
            events = response.text
            logging.info(f"Fetched events for category {category}")
            events_list.append((category, events))
        else:
            logging.error(f"Error fetching events for category {category}: {response.status_code}")
    return events_list

def structurize_csv_events(csv_string):
    """
    Parses CSV string into a pandas DataFrame, ensuring each event is a row.
    """
    # Check if the CSV string has multiple lines and commas
    lines = csv_string.strip().splitlines()
    if len(lines) > 1 and "," in lines[0]:
        # Likely standard CSV, use pandas
        df = pd.read_csv(StringIO(csv_string))
    else:
        # If not standard, try to parse manually (fallback)
        rows = [line.split(",") for line in lines if line]
        df = pd.DataFrame(rows[1:], columns=rows[0]) if rows else pd.DataFrame()
    return df

def save_events_to_excel(race_events, output_file):
    """
    Save the fetched race events to an Excel file.
    Each event will be its own row.
    """
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for category, events in race_events:
            try:
                df = structurize_csv_events(events)
                # Only save if df is not empty
                if not df.empty:
                    # Optionally, sort by id if present
                    if "id" in df.columns:
                        df = df.sort_values(by="id")
                    df.to_excel(writer, sheet_name=category, index=False)
                else:
                    logging.warning(f"No events to save for category {category}")
            except Exception as e:
                logging.error(f"Error processing events for category {category}: {e}")
    logging.info(f"Race events have been saved to {output_file}")

def main():
    race_categories = ["RACE_A", "RACE_B", "RACE_C"]
    race_events = get_race_events(athlete_id, username, api_key, race_categories, oldest_date, newest_date)
    if race_events:
        output_file = RACE_file_path
        save_events_to_excel(race_events, output_file)
        print(f"Race events have been saved to {output_file}")
    else:
        print("No RACE events found.")

if __name__ == "__main__":
    main()
