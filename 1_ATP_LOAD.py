import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

Athlete_TLA = "RAA" #Three letter Acronym of athlete.
ATP_year = "2026"
ATP_sheet_name = "ATP_Data"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\ATP2intervals_{Athlete_TLA}_{ATP_year}.xlsm"
parse_delay = .00
change_whole_range = True  # Variable to control whether to change the whole range or only upcoming targets

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

def format_activity_name(activity):
    if activity.lower() == 'mountainbikeride':
        return 'MountainBikeRide'
    if activity.lower() == 'openwaterswim':
        return 'OpenWaterSwim'
    if activity.lower() == 'gravelride':
        return 'GravelRide'
    if activity.lower() == 'trailrun':
        return 'TrailRun'
    return ''.join(word.capitalize() for word in activity.split('_'))

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
    for idx, row in df.iterrows():
        start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
        for col in df.columns:
            if col.endswith('_load_target'):
                activity = format_activity_name(col.split('_load')[0])
                load = int(row[col])
                time_col = f"{col.split('_load')[0]}_time"
                dist_col = f"{col.split('_load')[0]}_distance"
                time = int(row[time_col]) if time_col in row else 0
                distance = int(row[dist_col]) if dist_col in row else 0
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

    oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
    newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

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
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], errors='coerce')
    df = df.dropna(subset=['start_date_local'])
    efficient_event_sync(df, athlete_id, username, api_key)

if __name__ == "__main__":
    main()
