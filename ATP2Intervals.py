import logging
import os
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# User variables
excel_file_path = os.getenv('EXCEL_FILE_PATH', r"C:\TEMP\ATP.xlsx")  # Replace this with the location of your Excel file
sheet_name = os.getenv('SHEET_NAME', "ATP")  # Replace this with the name of the sheet
athlete_id = os.getenv('ATHLETE_ID', "athleteid")  # Replace this with your athlete_id
username = "API_KEY"  # This is always "API_KEY"
api_key = os.getenv('API_KEY', "yourapikey")  # Replace this with your API key
default_activity_type = os.getenv('DEFAULT_ACTIVITY_TYPE', "Bike")  # Default activity type
unit_preference = os.getenv('UNIT_PREFERENCE', "metric")  # User preference for units, default to metric

# Conversion factors
CONVERSION_FACTORS = {
    "metric": 1000,
    "imperial": 1609.344
}

# API endpoints
url_post = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events"
url_get = f"https://intervals.icu/api/v1/athlete/{athlete_id}/eventsjson"
url_delete = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events"
headers = {"Content-Type": "application/json"}

def create_update_or_delete_event(start_date, load_target, time_target, distance_target, activity_type, description, events):
    load_target = load_target or 0
    time_target = time_target or 0
    distance_target = distance_target or 0

    # Convert distance target based on unit preference for Bike and Run only
    if activity_type in ["Ride", "Run"]:
        distance_target *= CONVERSION_FACTORS[unit_preference]

    post_data = {
        "load_target": load_target,
        "time_target": time_target,
        "distance_target": distance_target,
        "category": "TARGET",
        "type": activity_type,
        "name": "Weekly",
        "start_date_local": start_date,
        "description": description
    }

    duplicate_event = next((event for event in events if event['category'] == "TARGET" and event['name'] == post_data['name'] and event['start_date_local'] == post_data['start_date_local'] and event['type'] == post_data['type']), None)

    if duplicate_event:
        event_id = duplicate_event['id']
        server_load_target = duplicate_event.get('load_target', 0) or 0
        server_time_target = duplicate_event.get('time_target', 0) or 0
        server_distance_target = duplicate_event.get('distance_target', 0) or 0

        if load_target == 0 and time_target == 0 and distance_target == 0 and not description:
            url_delete_event = f"{url_delete}/{event_id}"
            response_delete = requests.delete(url_delete_event, headers=headers, auth=HTTPBasicAuth(username, api_key))
            logging.info(f"DELETE Response Status Code: {response_delete.status_code}")
            if response_delete.status_code == 200:
                logging.info(f"Event deleted for {activity_type} on {start_date}!")
            else:
                logging.error(f"Error deleting event for {activity_type} on {start_date}: {response_delete.status_code}")
        elif (server_load_target != load_target or server_time_target != time_target or server_distance_target != distance_target or duplicate_event.get('description') != description):
            url_put = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{event_id}"
            put_data = {
                "load_target": load_target,
                "time_target": time_target,
                "distance_target": distance_target,
                "description": description
            }
            logging.info(f"Updating event: ID={event_id}, Data={put_data}")
            response_put = requests.put(url_put, headers=headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
            logging.info(f"PUT Response Status Code: {response_put.status_code}")
            if response_put.status_code == 200:
                logging.info(f"Duplicate event updated for {activity_type} on {start_date}!")
            else:
                logging.error(f"Error updating duplicate event for {activity_type} on {start_date}: {response_put.status_code}")
        else:
            logging.info(f"No changes needed for {activity_type} on {start_date}.")
    else:
        if load_target > 0 or time_target > 0 or distance_target > 0 or description:
            logging.info(f"New event: Data={post_data}")
            response_post = requests.post(url_post, headers=headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
            if response_post.status_code == 200:
                logging.info(f"New event created for {activity_type} on {start_date}!")
            else:
                logging.error(f"Error creating event for {activity_type} on {start_date}: {response_post.status_code}")

# Read the Excel file and specify the sheet
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
df.fillna(0, inplace=True)

# Get the oldest and newest date from the Excel list
oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

# Retrieve existing events
params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET", "resolve": "false"}
response_get = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(username, api_key))
events = response_get.json() if response_get.status_code == 200 else []

# Track if description has been added for the week
description_added = {}

for index, row in df.iterrows():
    start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
    swim_load, bike_load, run_load = int(row['swim_load']), int(row['bike_load']), int(row['run_load'])
    swim_time, bike_time, run_time = int(row['swim_time']), int(row['bike_time']), int(row['run_time'])
    swim_distance, bike_distance, run_distance = int(row['swim_distance']), int(row['bike_distance']), int(row['run_distance'])
    period = row['period'] if not pd.isna(row['period']) else ""
    focus = row['focus'] if not pd.isna(row['focus']) else ""
    week = row['start_date_local'].isocalendar()[1]
    description = f"You are in the {period} period." if period else ""
    if period == "Rest":
        description += " Stay in bed or on the beach!"
    if focus:
        description += f" Focus this week on {focus}."

    if week not in description_added:
        description_added[week] = False

    if swim_load > 0 or swim_time > 0 or swim_distance > 0:
        create_update_or_delete_event(start_date, swim_load, swim_time, swim_distance, "Swim", "", events)
    elif any([event['type'] == "Swim" for event in events if event['start_date_local'] == start_date]):
        create_update_or_delete_event(start_date, swim_load, swim_time, swim_distance, "Swim", "", events)
    if bike_load > 0 or bike_time > 0 or bike_distance > 0 or (description and not description_added[week]):
        create_update_or_delete_event(start_date, bike_load, bike_time, bike_distance, "Ride", description, events)
        description_added[week] = True
    elif any([event['type'] == "Ride" for event in events if event['start_date_local'] == start_date]):
        create_update_or_delete_event(start_date, bike_load, bike_time, bike_distance, "Ride", description, events)
    if run_load > 0 or run_time > 0 or run_distance > 0:
        create_update_or_delete_event(start_date, run_load, run_time, run_distance, "Run", "", events)
    elif any([event['type'] == "Run" for event in events if event['start_date_local'] == start_date]):
        create_update_or_delete_event(start_date, run_load, run_time, run_distance, "Run", "", events)
    if swim_load == 0 and bike_load == 0 and run_load == 0 and swim_time == 0 and bike_time == 0 and run_time == 0 and swim_distance == 0 and bike_distance == 0 and run_distance == 0:
        if period:
            create_update_or_delete_event(start_date, 0, 0, 0, default_activity_type, description, events)
        else:
            create_update_or_delete_event(start_date, 0, 0, 0, "Swim", "", events)
            create_update_or_delete_event(start_date, 0, 0, 0, "Ride", "", events)
            create_update_or_delete_event(start_date, 0, 0, 0, "Run", "", events)
