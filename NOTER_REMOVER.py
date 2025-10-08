import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

Athlete_TLA = "RAA"
ATP_year = "2026"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\ATP2intervals_{Athlete_TLA}_{ATP_year}.xlsm"
parse_delay = .001

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")
unit_preference = user_data.get('DISTANCE_SYSTEM', "metric")
note_ATP_color = user_data.get('NOTE_ATP_COLOR', "red")
do_at_rest = user_data.get('Do_At_Rest', "Do nothing!")
# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# API endpoints
url_base = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
API_headers = {"Content-Type": "application/json"}

#----------------------------------

#rip_word = 'weekly' or "Weekly" or "update" A word in the description of the note to be deleted.
rip_word = 'period' or 'Period'


# Function to delete events within a specified date range and category
def delete_weekly_note_events(athlete_id, username, api_key):
    start_date = datetime(2024, 1, 1).strftime("%Y-%m-%dT00:00:00")
    end_date = datetime(2026, 12, 31).strftime("%Y-%m-%dT23:59:59")
    url_get = f"{url_base}/events.json"
    params = {"oldest": start_date, "newest": end_date, "category": "NOTE"}
    logging.info(f"Fetching NOTE_EVENTS between {start_date} and {end_date}")
    response_get = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response_get.status_code != 200:
        logging.error(f"Error fetching events: {response_get.status_code}")
        return

    events = response_get.json()
    logging.info(f"Fetched {len(events)} NOTE_EVENTS")
    for event in events:
        if rip_word in event['name'].lower():
            event_id = event['id']
            logging.info(f"Deleting event ID={event_id} with name={event['name']}")
            url_del = f"{url_base}/events/{event_id}"
            response_del = requests.delete(url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
            if response_del.status_code == 200:
                logging.info(f"Deleted note event ID={event_id}")
            else:
                logging.error(f"Error deleting note event ID={event_id}: {response_del.status_code}")

# Main function to execute the script logic
def main():
    delete_weekly_note_events(athlete_id, username, api_key)

if __name__ == "__main__":
    main()
