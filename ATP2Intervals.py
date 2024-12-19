import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime

# User variables
excel_file_path = r"C:\TEMP\ATP.xlsx"  # Replace this with the location of your Excel file
sheet_name = "ATP"  # Replace this with the name of the sheet
athlete_id = "your_athlete_ID"  # Replace this with your athlete_id
username = "API_KEY"  # This is always "API_KEY" (not quite a variable....now)
api_key = "yourapikey"  # Replace this with your API key

# API endpoints
url_post = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events"
url_get = f"https://intervals.icu/api/v1/athlete/{athlete_id}/eventsjson"
url_delete = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events"
headers = {
    "Content-Type": "application/json"
}

# Function to create, update, or delete events
def create_update_or_delete_event(start_date, load_target, time_target, distance_target, activity_type, description, events):
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
        if load_target == 0 and time_target == 0 and distance_target == 0:
            url_delete_event = f"{url_delete}/{event_id}"
            response_delete = requests.delete(url_delete_event, headers=headers, auth=HTTPBasicAuth(username, api_key))
            print(f"DELETE Response Status Code: {response_delete.status_code}")
            if response_delete.status_code == 200:
                print(f"Event deleted for {activity_type} on {start_date}!")
            else:
                print(f"Error deleting event for {activity_type} on {start_date}: {response_delete.status_code}")
        elif (duplicate_event['load_target'] != load_target or 
              duplicate_event['time_target'] != time_target or 
              duplicate_event['distance_target'] != distance_target or 
              duplicate_event['description'] != description):
            url_put = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{event_id}"
            put_data = {
                "load_target": load_target,
                "time_target": time_target,
                "distance_target": distance_target,
                "description": description
            }
            print(f"Updating event: ID={event_id}, Data={put_data}")
            response_put = requests.put(url_put, headers=headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
            print("PUT Response Status Code:", response_put.status_code)
            if response_put.status_code == 200:
                print(f"Duplicate event updated for {activity_type} on {start_date}!")
            else:
                print(f"Error updating duplicate event for {activity_type} on {start_date}: {response_put.status_code}")
        else:
            print(f"No changes needed for {activity_type} on {start_date}.")
    else:
        if load_target > 0 or time_target > 0 or distance_target > 0 or description:
            print(f"New event: Data={post_data}")
            response_post = requests.post(url_post, headers=headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
            if response_post.status_code == 200:
                print(f"New event created for {activity_type} on {start_date}!")
            else:
                print(f"Error creating event for {activity_type} on {start_date}: {response_post.status_code}")

# Read the Excel file and specify the sheet
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Fill NaN values with 0
df.fillna(0, inplace=True)

# Get the oldest and newest date from the Excel list
oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

# Parameters for retrieving events with category "TARGET" within the full date range
params = {
    "oldest": oldest_date,
    "newest": newest_date,
    "category": "TARGET",
    "resolve": "false"
}

# Send the GET request to retrieve existing events
response_get = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(username, api_key))
if response_get.status_code == 200:
    try:
        events = response_get.json()
    except ValueError:
        print("Error decoding JSON response.")
        events = []
else:
    print(f"Error retrieving events: {response_get.status_code}")
    events = []

# Track if description has been added for the week
description_added = {}

for index, row in df.iterrows():
    try:
        start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
    except KeyError:
        print("Column 'start_date_local' not found in the Excel file.")
        continue

    swim_load = int(row['swim_load']) if not pd.isna(row['swim_load']) else 0
    bike_load = int(row['bike_load']) if not pd.isna(row['bike_load']) else 0
    run_load = int(row['run_load']) if not pd.isna(row['run_load']) else 0
    swim_time = int(row['swim_time']) if not pd.isna(row['swim_time']) else 0
    bike_time = int(row['bike_time']) if not pd.isna(row['bike_time']) else 0
    run_time = int(row['run_time']) if not pd.isna(row['run_time']) else 0
    swim_distance = int(row['swim_distance']) if not pd.isna(row['swim_distance']) else 0
    bike_distance = int(row['bike_distance']) if not pd.isna(row['bike_distance']) else 0
    run_distance = int(row['run_distance']) if not pd.isna(row['run_distance']) else 0
    period = row['period']
    week = row['start_date_local'].isocalendar()[1]
    description = f"You are in the {period} period."

    if week not in description_added:
        description_added[week] = False

    if swim_load > 0 or swim_time > 0 or swim_distance > 0 or (description and not description_added[week]):
        create_update_or_delete_event(start_date, swim_load, swim_time, swim_distance, "Swim", description if not description_added[week] else "", events)
        description_added[week] = True
    if bike_load > 0 or bike_time > 0 or bike_distance > 0 or (description and not description_added[week]):
        create_update_or_delete_event(start_date, bike_load, bike_time, bike_distance, "Ride", description if not description_added[week] else "", events)
        description_added[week] = True
    if run_load > 0 or run_time > 0 or run_distance > 0 or (description and not description_added[week]):
        create_update_or_delete_event(start_date, run_load, run_time, run_distance, "Run", description if not description_added[week] else "", events)
        description_added[week] = True
