import logging
import os
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
#from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def format_activity_name(activity):
    return ''.join(word.capitalize() for word in activity.split('_'))

# Function to read user data from USERDATA.xlsx file
def read_user_data(excel_file_path, sheet_name="userdata"):
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

# Read user data from USERDATA.xlsx file
user_data = read_user_data(r'C:\Temp\USERDATA.xlsx')
excel_file_path = user_data.get('EXCEL_FILE_PATH', r"C:\TEMP\ATP.xlsx")
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")

# User variables
sheet_name = os.getenv('SHEET_NAME', "ATP")  # Replace this with the name of the sheet
default_activity_type = os.getenv('DEFAULT_ACTIVITY_TYPE', "Bike")  # Default activity type
unit_preference = os.getenv('UNIT_PREFERENCE', "metric")  # User preference for units, default to metric
note_color = "red"
whattodowithrest = "**Stay in bed or on the beach! :-)**"

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

def create_update_or_delete_note_event(start_date, description, events):
    end_date = start_date  # End date is the same as start date for NOTE events
    
    post_data = {
        "category": "NOTE",
        "start_date_local": start_date,
        "end_date_local": end_date,
        "name": "Weekly Summary",
        "description": description,
        "not_on_fitness_chart": "true",
        "show_as_note": "false",
        "show_on_ctl_line": "false",
        "athlete_cannot_edit": "false",
        "color": note_color
    }

    duplicate_event = next((event for event in events if event['category'] == "NOTE" and event['name'] == post_data['name'] and event['start_date_local'] == post_data['start_date_local']), None)
    
    if duplicate_event:
        event_id = duplicate_event['id']
        if duplicate_event.get('description') != description:
            url_put = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{event_id}"
            put_data = {
                "description": description
            }
            logging.info(f"Updating event: ID={event_id}, Data={put_data}")
            response_put = requests.put(url_put, headers=headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
            logging.info(f"PUT Response Status Code: {response_put.status_code}")
            if response_put.status_code == 200:
                logging.info(f"Duplicate event updated on {start_date}!")
            else:
                logging.error(f"Error updating duplicate event on {start_date}: {response_put.status_code}")
        else:
            logging.info(f"No changes needed for event on {start_date}.")
    else:
        logging.info(f"New event: Data={post_data}")
        response_post = requests.post(url_post, headers=headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
        if response_post.status_code == 200:
            logging.info(f"New event created on {start_date}!")
        else:
            logging.error(f"Error creating event on {start_date}: {response_post.status_code}")

# Add this function definition to ATP2Intervals.py


def create_update_or_delete_target_event(start_date, load_target, time_target, distance_target, activity_type, description, events):
    load_target = load_target or 0
    time_target = time_target or 0
    distance_target = distance_target or 0

    # Convert distance target based on unit preference for Ride and Run only
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

    duplicate_event = next((event for event in events if event['category'] == "TARGET" and event['name'] == post_data['name'] and event['start_date_local'] == post_data['start_date_local'] and event['type'] == activity_type), None)

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
params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET,NOTE", "resolve": "false"}
response_get = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(username, api_key))
events = response_get.json() if response_get.status_code == 200 else []

# Track if description has been added for the week
description_added = {}

focus_columns = [
    'Aerobic Endurance', 'Muscular force', 'Speed Skills',
    'Muscular Endurance', 'Anaerobic Endurance', 'Sprint Power'
]

def format_focus_items(focus_items):
    if len(focus_items) > 1:
        return ', '.join(focus_items[:-1]) + ' and ' + focus_items[-1]
    return ''.join(focus_items)

# First handle all TARGET events
for index, row in df.iterrows():
    start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
    period = row['period'] if not pd.isna(row['period']) else ""
    focus = row['focus'] if not pd.isna(row['focus']) else ""
    test = row['test'] if not pd.isna(row['test']) else ""  # Added test column
    week = row['start_date_local'].isocalendar()[1]
    description = f"You are in the {period} period.\n\n" if period else ""
    if period == "Rest":
        description += whattodowithrest
    if test:  # Add test comment if there is a value
        description += f"Do the following test(s) this week: **{test}**.\n\n"
    if focus:
        description += f"Focus this week on {focus}.\n\n"
    
    # Add focus based on specified columns
    additional_focus = [col for col in focus_columns if str(row.get(col, '')).lower() == 'x']
    if additional_focus:
        formatted_focus = format_focus_items(additional_focus)
        description += f"Focus on {formatted_focus}.\n\n"
    
    # Add focus for A, B, and C category races
    race_cat = str(row.get('cat', '')).upper()
    race_name = row.get('race', '').strip()
    if race_cat == 'A' and race_name:
        description += f"Use the **{race_name}** as an {race_cat}-event to primarily focus this week on this race."
    elif race_cat == 'B' and race_name:
        description += f"Use the **{race_name}** to learn and improve skills."
    elif race_cat == 'C' and race_name:
        description += f"Use the **{race_name}** as hard effort training or just having fun!"

    if week not in description_added:
        description_added[week] = False

    for col in df.columns:
        if col.endswith('_load'):
            activity = format_activity_name(col.split('_load')[0])
            load = int(row[col])
            time_col = f"{col.split('_load')[0]}_time"
            distance_col = f"{col.split('_load')[0]}_distance"
            time = int(row[time_col]) if time_col in row else 0
            distance = int(row[distance_col]) if distance_col in row else 0

            if load > 0 or time > 0 or distance > 0:
                create_update_or_delete_target_event(start_date, load, time, distance, activity, description if activity == "Ride" and not description_added[week] else "", events)
                if activity == "Ride" and not description_added[week]:
                    description_added[week] = True
            elif any([event['type'] == activity for event in events if event['start_date_local'] == start_date]):
                create_update_or_delete_target_event(start_date, load, time, distance, activity, "", events)

# Then handle NOTE events
for index, row in df.iterrows():
    start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
    period = row['period'] if not pd.isna(row['period']) else ""
    focus = row['focus'] if not pd.isna(row['focus']) else ""
    test = row['test'] if not pd.isna(row['test']) else ""  # Added test column
    week = row['start_date_local'].isocalendar()[1]
    description = f"You are in the {period} period.\n\n" if period else ""
    if period == "Rest":
        description += whattodowithrest
    if test:  # Add test comment if there is a value
        description += f"Do the following test(s) this week: **{test}**.\n\n"
    if focus:
        description += f"Focus this week on {focus}.\n\n"
    
    # Add focus based on specified columns
    additional_focus = [col for col in focus_columns if str(row.get(col, '')).lower() == 'x']
    if additional_focus:
        formatted_focus = format_focus_items(additional_focus)
        description += f"Focus on {formatted_focus}.\n\n"
    
    # Add focus for A, B, and C category races
    race_cat = str(row.get('cat', '')).upper()
    race_name = row.get('race', '').strip()
    if race_cat == 'A' and race_name:
        description += f"Use the **{race_name}** as an {race_cat}-event to primarily focus this week on this race."
    elif race_cat == 'B' and race_name:
        description += f"Use the **{race_name}** to learn and improve skills."
    elif race_cat == 'C' and race_name:
        description += f"Use the **{race_name}** as hard effort training or just having fun!"

    if week not in description_added:
        description_added[week] = False

    # Handle NOTE events
    if all(row[col] == 0 for col in df.columns if col.endswith('_load')) and all(row[col] == 0 for col in df.columns if col.endswith('_time')) and all(row[col] == 0 for col in df.columns if col.endswith('_distance')):
        if period:
            create_update_or_delete_note_event(start_date, description, events)
        else:
            for col in df.columns:
                if col.endswith('_load'):
                    activity = format_activity_name(col.split('_load')[0])
                    create_update_or_delete_note_event(start_date, "", events)
