import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(level)s - %(message)s')
#logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

Athlete_TLA = "TLA" #Three letter Acronym of athlete.
ATP_sheet_name = "ATP_Data"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\Intervals_API_Tools_Office365_v1.6_ATP2intervals_{Athlete_TLA}.xlsm"
parse_delay = .01

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
    if activity.lower() == 'trailrun':
        return 'TrailRun'
    return ''.join(word.capitalize() for word in activity.split('_'))

def distance_conversion_factor(unit_preference):
    conversion_factors = {
        "metric": 1000,
        "imperial": 1609.344
    }
    return conversion_factors.get(unit_preference, 1000)

def delete_events(athlete_id, username, api_key, oldest_date, newest_date, category, name=None):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": category}
    response_get = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    for event in events:
        if name and event['name'] != name:
            continue
        event_id = event['id']
        url_del = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
        response_del = requests.delete(url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
        if response_del.status_code == 200:
            logging.info(f"Deleted {category.lower()} event ID={event_id}")
        else:
            logging.error(f"Error deleting {category.lower()} event ID={event_id}: {response_del.status_code}")
        time_module.sleep(parse_delay)

def create_update_or_delete_target_event(start_date, load_target, time_target, distance_target, activity_type, events, athlete_id, username, api_key):
    if load_target is None or load_target == 0:
        logging.info(f"Skipping {activity_type} event on {start_date} due to None or 0 load target.")
        return

    load_target = load_target or 0
    time_target = time_target or 0
    distance_target = distance_target or 0

    if activity_type not in ["Swim" , "OpenwaterSwim"]: #all sporttypes that are given in meters/yards instead of km/mi.
        distance_target *= distance_conversion_factor(unit_preference)

    post_data = {
        "load_target": load_target,
        "time_target": time_target,
        "distance_target": distance_target,
        "category": "TARGET",
        "type": activity_type,
        "name": "Weekly",
        "start_date_local": start_date
    }

    duplicate_event = next((event for event in events if event['category'] == "TARGET" and event['name'] == post_data['name'] and event['start_date_local'] == post_data['start_date_local']), None)

    if duplicate_event:
        event_id = duplicate_event['id']
        server_load_target = duplicate_event.get('load_target', 0) or 0
        server_time_target = duplicate_event.get('time_target', 0) or 0
        server_distance_target = duplicate_event.get('distance_target', 0) or 0

        if server_load_target != load_target or server_time_target != time_target or server_distance_target != distance_target:
            url_put = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
            put_data = {
                "load_target": load_target,
                "time_target": time_target,
                "distance_target": distance_target
            }
            logging.info(f"Updating event: ID={event_id}, Data={put_data}")
            response_put = requests.put(url_put, headers=API_headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
            logging.info(f"PUT Response Status Code: {response_put.status_code}")
            if response_put.status_code == 200:
                logging.info(f"Duplicate event updated for {activity_type} on {start_date}!")
            else:
                logging.error(f"Error updating duplicate event for {activity_type} on {start_date}: {response_put.status_code}")
        else:
            logging.info(f"No changes needed for {activity_type} on {start_date}.")
    else:
        if load_target > 0 or time_target > 0 or distance_target > 0:
            logging.info(f"New event: Data={post_data}")
            url_post = f"{url_base}/events".format(athlete_id=athlete_id)
            response_post = requests.post(url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
            if response_post.status_code == 200:
                logging.info(f"New event created for {activity_type} on {start_date}!")
            else:
                logging.error(f"Error creating event for {activity_type} on {start_date}: {response_post.status_code}")
            time_module.sleep(parse_delay)

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)

    oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
    newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

    delete_events(athlete_id, username, api_key, oldest_date, newest_date, "TARGET")

    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET,NOTE", "resolve": "false"}
    response_get = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    for index, row in df.iterrows():
        start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
        for col in df.columns:
            if col.endswith('_load_target'):
                activity = format_activity_name(col.split('_load')[0])
                load = int(row[col])
                time_col = f"{col.split('_load')[0]}_time"
                distance_col = f"{col.split('_load')[0]}_distance"
                time = int(row[time_col]) if time_col in row else 0
                distance = int(row[distance_col]) if distance_col in row else 0

                if load > 0 or time > 0 or distance > 0:
                    create_update_or_delete_target_event(start_date, load, time, distance, activity, events, athlete_id, username, api_key)
                elif any([event['type'] == activity for event in events if event['start_date_local'] == start_date]):
                    create_update_or_delete_target_event(start_date, load, time, distance, activity, events, athlete_id, username, api_key)

        if all(row[col] == 0 for col in df.columns if col.endswith('_load')) and all(row[col] == 0 for col in df.columns if col.endswith('_time')) and all(row[col] == 0 for col in df.columns if col.endswith('_distance')):
            for col in df.columns:
                if col.endswith('_load'):
                    activity = format_activity_name(col.split('_load')[0])
                    create_update_or_delete_target_event(start_date, 0, 0, 0, activity, events, athlete_id, username, api_key)
            time_module.sleep(parse_delay)

if __name__ == "__main__":
    main()
