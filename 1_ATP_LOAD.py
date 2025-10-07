import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta

Athlete_TLA = "TLA" #Three letter Acronym of athlete.
ATP_year = "YYYY"
ATP_sheet_name = "ATP_Data"
ATP_sheet_Conditions = "ATP_Conditions"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\ATP2intervals_{Athlete_TLA}_{ATP_year}.xlsm"
parse_delay = .00
change_whole_range = True  # Variable to control whether to change the whole range or only upcoming targets

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

def parse_atp_date(date_str):
    # Try common formats
    for fmt in ("%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(date_str), fmt)
        except ValueError:
            continue
    raise ValueError(f"Date '{date_str}' is not in a recognized format.")

def read_ATP_period(ATP_file_path, sheet_name=ATP_sheet_Conditions):
    df_cond = pd.read_excel(ATP_file_path, sheet_name=sheet_name, usecols="B:C")
    cond_dict = dict(zip(df_cond.iloc[:, 0], df_cond.iloc[:, 1]))
    start_str = cond_dict.get("Start_ATP")
    end_str = cond_dict.get("End_ATP")
    start_date = parse_atp_date(start_str)
    end_date = parse_atp_date(end_str)
    oldest_date = start_date.strftime("%Y-%m-%dT00:00:00")
    newest_date = end_date.strftime("%Y-%m-%dT00:00:00")
    # Return only oldest_date and newest_date for your usage
    return oldest_date, newest_date

def prompt_overwrite_past():
    answer = input("Do you want to overwrite data in the past? (yes/no): ").strip().lower()
    return answer == "yes"

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")
unit_preference = user_data.get('DISTANCE_SYSTEM', "metric")
note_ATP_color = user_data.get('NOTE_ATP_COLOR', "red")
do_at_rest = user_data.get('Do_At_Rest', "Do nothing!")

url_base = "https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
url_activities = f"https://intervals.icu/api/v1/athlete/{athlete_id}/activities"
API_headers = {"Content-Type": "application/json"}

def clean_activity_name(col_name):
    # Strip '_load' or '_load_target' from column names to get the activity name
    return col_name.replace('_load_target', '').replace('_load', '')

def distance_conversion_factor(unit_preference):
    conversion_factors = {
        "metric": 1000,
        "imperial": 1609.344
    }
    return conversion_factors.get(unit_preference, 1000)

def get_existing_events(athlete_id, oldest_date, newest_date, username, api_key):
    url_get = f"https://intervals.icu/api/v1/athlete/{athlete_id}/eventsjson"
    params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET"}
    response = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        events = response.json()
        event_map = {
            (e['start_date_local'], e['type']): e
            for e in events
        }
        return event_map
    else:
        logging.error(f"Failed to fetch events ({response.status_code})")
        return {}

def get_desired_events(df):
    desired = {}
    dist_factor = distance_conversion_factor(unit_preference)
    for idx, row in df.iterrows():
        start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
        for col in df.columns:
            if col.endswith('_load_target'):
                activity = clean_activity_name(col)
                load = int(row[col])
                time_col = f"{activity}_time"
                dist_col = f"{activity}_distance"
                time = int(row[time_col]) * 60 if time_col in row else 0  # Convert seconds to minutes
                # Only convert distance for non-swimming activities
                if dist_col in row:
                    if activity.lower() in ['swim', 'openwaterswim']:
                        distance = int(row[dist_col])  # Swimming stays meters
                    else:
                        distance = int(row[dist_col]) * dist_factor  # Convert for non-swimming
                else:
                    distance = 0
                key = (start_date, activity)
                desired[key] = {
                    'start_date_local': start_date,
                    'type': activity,
                    'load_target': load,
                    'time_target': time,
                    'distance_target': distance
                }
    return desired

def efficient_event_sync(df, athlete_id, username, api_key):
    if df.empty:
        logging.error("No valid dates found in 'start_date_local'.")
        return

    # Use ATP period from ATP_Conditions
    oldest_date, newest_date = read_ATP_period(ATP_file_path)

    # Get existing events from the server
    existing_events = get_existing_events(athlete_id, oldest_date, newest_date, username, api_key)

    # Build desired events from the DataFrame
    desired_events = get_desired_events(df)

    # 1. Create or Update events
    for key, new_event in desired_events.items():
        old_event = existing_events.get(key)
        if old_event:
            # Only update if something changed
            if (
                old_event.get('load_target', 0) != new_event['load_target'] or
                old_event.get('time_target', 0) != new_event['time_target'] or
                old_event.get('distance_target', 0) != new_event['distance_target']
            ):
                url_put = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{old_event['id']}"
                put_data = {
                    "load_target": new_event['load_target'],
                    "time_target": new_event['time_target'],
                    "distance_target": new_event['distance_target']
                }
                logging.info(f"Updating event {key}: {put_data}")
                response_put = requests.put(url_put, headers=API_headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
                if response_put.status_code == 200:
                    logging.info(f"Updated event for {key}")
                else:
                    logging.error(f"Failed to update event for {key}: {response_put.status_code}")
                time_module.sleep(parse_delay)
            else:
                logging.info(f"No changes needed for event {key}")
        else:
            # Create new event
            if any([new_event['load_target'] > 0, new_event['time_target'] > 0, new_event['distance_target'] > 0]):
                url_post = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events"
                post_data = {
                    "load_target": new_event['load_target'],
                    "time_target": new_event['time_target'],
                    "distance_target": new_event['distance_target'],
                    "category": "TARGET",
                    "type": new_event['type'],
                    "name": "Weekly",
                    "start_date_local": new_event['start_date_local']
                }
                logging.info(f"Creating event {key}: {post_data}")
                response_post = requests.post(url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
                if response_post.status_code == 200:
                    logging.info(f"Created new event for {key}")
                else:
                    logging.error(f"Failed to create event for {key}: {response_post.status_code}")
                time_module.sleep(parse_delay)

    # 2. Delete events that are no longer needed
    for key, old_event in existing_events.items():
        if key not in desired_events:
            url_del = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{old_event['id']}"
            logging.info(f"Deleting event {key}")
            response_del = requests.delete(url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
            if response_del.status_code == 200:
                logging.info(f"Deleted event for {key}")
            else:
                logging.error(f"Failed to delete event for {key}: {response_del.status_code}")
            time_module.sleep(parse_delay)

def main():
    # Prompt user for overwriting past data
    overwrite_past = prompt_overwrite_past()
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], errors='coerce')
    df = df.dropna(subset=['start_date_local'])

    # Strictly limit data to ATP period
    oldest_date, newest_date = read_ATP_period(ATP_file_path)
    oldest = pd.to_datetime(oldest_date)
    newest = pd.to_datetime(newest_date)
    df = df[(df['start_date_local'] >= oldest) & (df['start_date_local'] <= newest)]

    if not overwrite_past:
        # Only keep events with start_date_local in the future (relative to now)
        now = datetime.now()
        df = df[df['start_date_local'] >= now]

    efficient_event_sync(df, athlete_id, username, api_key)

if __name__ == "__main__":
    main()


