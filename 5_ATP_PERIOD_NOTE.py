import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

Athlete_TLA = "TLA" # Three letter Acronym of athlete.
ATP_sheet_name = "ATP_Data"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\Intervals_API_Tools_Office365_v1.6_ATP2intervals_{Athlete_TLA}.xlsm"

parse_delay = .01
note_PERIOD_name = 'Period:'

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")

url_base = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
url_activities = f"https://intervals.icu/api/v1/athlete/{athlete_id}/activities"
API_headers = {"Content-Type": "application/json"}

def get_last_day_of_week(date):
    return date + timedelta(days=(6 - date.weekday()))

def get_period_end_date(df, start_index):
    period = df.at[start_index, 'period']
    for i in range(start_index, len(df)):
        if df.at[i, 'period'] != period:
            return get_last_day_of_week(df.at[i-1, 'start_date_local'])
    return get_last_day_of_week(df.at[len(df)-1, 'start_date_local'])

def get_note_color(period):
    base_period = period.split()[0]  # Get the base period name (e.g., "Base" from "Base 1")
    color_mapping = {
        "Base": "yellow",
        "Peak": "orange",
        "Race": "red",
        "Trans": "green",
        "Prep": "blue",
        "Recovery": "purple",
        "Rest": "cyan",
        "Build": "blue"
    }
    return color_mapping.get(base_period, "black")  # Default to black if base period not found

def delete_events(athlete_id, username, api_key, oldest_date, newest_date, category, name_prefix):
    url_get = f"{url_base}/events".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": category}
    response_get = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    for event in events:
        if name_prefix and not event['name'].startswith(name_prefix):
            continue
        event_id = event['id']
        url_del = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
        response_del = requests.delete(url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
        if response_del.status_code == 200:
            logging.info(f"Deleted {category.lower()} event ID={event_id}")
        else:
            logging.error(f"Error deleting {category.lower()} event ID={event_id}: {response_del.status_code}")

def create_note_event(start_date, end_date, description, period, athlete_id, username, api_key):
    url_base = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
    url_post = f"{url_base}/events"
    API_headers = {"Content-Type": "application/json"}
    
    period_full = period
    if period == "Trans":
        period_full = "Transition"
    elif period == "Prep":
        period_full = "Preparation"
    
    color = get_note_color(period)
    
    post_data = {
        "category": "NOTE",
        "start_date_local": start_date.strftime("%Y-%m-%dT00:00:00"),
        "end_date_local": (end_date + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"),  # Add an extra day
        "name": f"{note_PERIOD_name} {period_full}",
        "description": description,
        "color": color
    }
    
    response_post = requests.post(url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
    if response_post.status_code == 200:
        logging.info(f"New event created from {start_date} to {end_date}!")
    else:
        logging.error(f"Error creating event: {response_post.status_code} - {response_post.text}")

def get_first_a_event(df, note_event_date):
    note_date = datetime.strptime(note_event_date, "%Y-%m-%dT00:00:00")
    for index, row in df.iterrows():
        event_date = pd.to_datetime(row.get('start_date_local'))
        if event_date > note_date and str(row.get('cat', '')).upper() == 'A' and row.get('race', '').strip():
            return row.get('race', '').strip()
    return None

def populate_race_description(description, first_a_event):
    if first_a_event:
        description = f"This (part) of the plan aims for **{first_a_event}**.\n\n" + description
    return description

def create_description(period, start_date, end_date, first_a_event):
    description = f"You are in the **{period}-period** (from {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')})"
    description = populate_race_description(description, first_a_event)
    return description

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name, engine='openpyxl')
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], format='%d-%b')
    
    # Explicitly cast columns to compatible dtype
    df['period'] = df['period'].astype(str)
    df['start_date_local'] = df['start_date_local'].astype('datetime64[ns]')
    df.fillna('', inplace=True)
    
    # Define date range for deleting events
    oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
    newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")
    
    # Delete existing period notes
    delete_events(athlete_id, username, api_key, oldest_date, newest_date, "NOTE", note_PERIOD_name)
    
    for i in range(len(df)):
        start_date = df.at[i, 'start_date_local']
        period = df.at[i, 'period']
        
        if i == 0 or df.at[i-1, 'period'] != period:
            end_date = get_period_end_date(df, i)
            first_a_event = get_first_a_event(df, start_date.strftime("%Y-%m-%dT00:00:00"))
            description = create_description(period, start_date, end_date, first_a_event)
            create_note_event(start_date, end_date, description, period, athlete_id, username, api_key)

if __name__ == "__main__":
    main()
