import logging
import os
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module  # Rename the time module to avoid conflict

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levellevel)s - %(message)s')

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
whattodowithrest = "**Stay in bed or on the beach! :-)**"
note_color = "red"
note_name = "Weekly Summary"

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

def delete_target_events(oldest_date, newest_date):
    # Retrieve existing events within the date range
    params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET"}
    response_get = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    # Delete TARGET events within the date range
    for event in events:
        if event['category'] == "TARGET":
            event_id = event['id']
            url_del = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{event_id}"
            response_del = requests.delete(url_del, headers=headers, auth=HTTPBasicAuth(username, api_key))
            if response_del.status_code == 200:
                logging.info(f"Deleted target event ID={event_id}")
            else:
                logging.error(f"Error deleting target event ID={event_id}: {response_del.status_code}")

def delete_note_events(oldest_date, newest_date):
    # Retrieve existing NOTE events within the date range
    params = {"oldest": oldest_date, "newest": newest_date, "category": "NOTE"}
    response_get = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    # Delete NOTE events within the date range with the name defined in note_name
    for event in events:
        if event['category'] == "NOTE" and event['name'] == note_name:
            event_id = event['id']
            url_del = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{event_id}"
            response_del = requests.delete(url_del, headers=headers, auth=HTTPBasicAuth(username, api_key))
            if response_del.status_code == 200:
                logging.info(f"Deleted note event ID={event_id}")
            else:
                logging.error(f"Error deleting note event ID={event_id}: {response_del.status_code}")

def create_update_or_delete_target_event(start_date, load_target, time_target, distance_target, activity_type, events):
    if load_target is None or load_target == 0:
        logging.info(f"Skipping {activity_type} event on {start_date} due to None or 0 load target.")
        return

    load_target = load_target or 0
    time_target = time_target or 0
    distance_target = distance_target or 0

    # Convert distance target based on unit preference for Ride and Run only
    if activity_type in ["Ride", "Run"]:
        distance_target *= CONVERSION_FACTORS["metric"]

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

        # Update the event if targets have changed
        if server_load_target != load_target or server_time_target != time_target or server_distance_target != distance_target:
            url_put = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{event_id}"
            put_data = {
                "load_target": load_target,
                "time_target": time_target,
                "distance_target": distance_target
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
        if load_target > 0 or time_target > 0 or distance_target > 0:
            logging.info(f"New event: Data={post_data}")
            response_post = requests.post(url_post, headers=headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
            if response_post.status_code == 200:
                logging.info(f"New event created for {activity_type} on {start_date}!")
            else:
                logging.error(f"Error creating event for {activity_type} on {start_date}: {response_post.status_code}")
            time_module.sleep(0.1)  # Add a 100ms delay between each add event

# Read the Excel file and specify the sheet
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
df.fillna(0, inplace=True)

# Get the oldest and newest date from the Excel list
oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

# Delete existing TARGET and NOTE events in the date range
delete_target_events(oldest_date, newest_date)
delete_note_events(oldest_date, newest_date)

# Retrieve existing events
params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET,NOTE", "resolve": "false"}
response_get = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(username, api_key))
events = response_get.json() if response_get.status_code == 200 else []

# Handle all TARGET events first
for index, row in df.iterrows():
    start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
    for col in df.columns:
        if col.endswith('_load'):
            activity = format_activity_name(col.split('_load')[0])
            load = int(row[col])
            time_col = f"{col.split('_load')[0]}_time"
            distance_col = f"{col.split('_load')[0]}_distance"
            time = int(row[time_col]) if time_col in row else 0
            distance = int(row[distance_col]) if distance_col in row else 0

            if load > 0 or time > 0 or distance > 0:
                create_update_or_delete_target_event(start_date, load, time, distance, activity, events)
            elif any([event['type'] == activity for event in events if event['start_date_local'] == start_date]):
                create_update_or_delete_target_event(start_date, load, time, distance, activity, events)

    if all(row[col] == 0 for col in df.columns if col.endswith('_load')) and all(row[col] == 0 for col in df.columns if col.endswith('_time')) and all(row[col] == 0 for col in df.columns if col.endswith('_distance')):
        for col in df.columns:
            if col.endswith('_load'):
                activity = format_activity_name(col.split('_load')[0])
                create_update_or_delete_target_event(start_date, 0, 0, 0, activity, events)

# Handle NOTE events (consolidated weekly notes)
def create_update_or_delete_note_event(start_date, description, color, events):
    end_date = start_date  # End date is the same as start date for NOTE events

    # Provide default description if empty
    if not description:
        description = "Nothing to mention this week."

    post_data = {
        "category": "NOTE",
        "start_date_local": start_date,
        "end_date_local": end_date,
        "name": note_name,
        "description": description,
        "not_on_fitness_chart": "true",
        "show_as_note": "false",
        "show_on_ctl_line": "false",
        "athlete_cannot_edit": "false",
        "color": note_color
    }

    logging.info(f"New event: Data={post_data}")
    response_post = requests.post(url_post, headers=headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
    if response_post.status_code == 200:
        logging.info(f"New event created on {start_date}!")
    else:
        logging.error(f"Error creating event on {start_date}: {response_post.status_code}")

# Track if description has been added for the week
description_added = {}

def get_athlete_profile(athlete_id, username, api_key):
    url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
    response = requests.get(url_profile, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        profile = response.json()
        first_name = profile.get('Name', 'Athlete ohne name')
        return first_name
    else:
        logging.error(f"Error fetching athlete profile: {response.status_code}")
        return None

# Fetch the athlete's first name and print a greeting
athlete_first_name = get_athlete_profile(athlete_id, username, api_key)
if athlete_first_name:
    print(f"Hello {athlete_first_name}!")

def format_focus_items_notes(focus_items_notes):
    if len(focus_items_notes) > 1:
        return ', '.join(focus_items_notes[:-1]) + ' and ' + focus_items_notes[-1]
    return ''.join(focus_items_notes)

for index, row in df.iterrows():
    start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
    period = row['period'] if not pd.isna(row['period']) else ""
    focus = row['focus'] if not pd.isna(row['focus']) else ""
    test = row['test'] if not pd.isna(row['test']) else ""  # Added test column
    week = row['start_date_local'].isocalendar()[1]
    description = f"- You are in the **{period}** period of your trainingplan.\n\n" if period else ""
    if period == "Rest":
        description += f"- {whattodowithrest}\n\n"
    if test:  # Add test comment if there is a value
        description += f"- Do the following test(s) this week: **{test}**.\n\n"
    if focus:
        description += f"- Focus this week on {focus}.\n\n"
    
    # Add focus based on specified columns
    focus_columns = [
        'Aerobic Endurance', 'Muscular force', 'Speed Skills',
        'Muscular Endurance', 'Anaerobic Endurance', 'Sprint Power'
    ]
    additional_focus = [col for col in focus_columns if str(row.get(col, '')).lower() == 'x']
    if additional_focus:
        formatted_focus = format_focus_items_notes(additional_focus)
        description += f"- Focus on {formatted_focus}.\n\n"
    elif description.strip():
        description += "- You don't have to focus on specific workouts this week.\n\n"
    
    # Add focus for A, B, and C category races
    race_cat = str(row.get('cat', '')).upper()
    race_name = row.get('race', '').strip()
    if race_cat == 'A' and race_name:
        description += f"- Use the **{race_name}** as an {race_cat}-event to primarily focus this week on this race.\n\n"
    elif race_cat == 'B' and race_name:
        description += f"- Use the **{race_name}** to learn and improve skills.\n\n"
    elif race_cat == 'C' and race_name:
        description += f"- Use the **{race_name}** as a hard effort training or just having fun!\n\n"
    
    # Add extra data about the next race
    next_race = None
    for i in range(index + 2, len(df)):
        next_race_name = df.at[i, 'race']
        if next_race_name and next_race_name not in ['-', '0', 'None']:
            next_race = df.iloc[i]
            break
    if next_race is not None:
        next_race_date = next_race['start_date_local'].strftime("%Y-%m-%dT00:00:00")
        next_race_cat = str(next_race.get('cat', '')).upper()
        next_race_name = next_race.get('race', '').strip()
        next_race_week = next_race['start_date_local'].isocalendar()[1]
        weeks_to_go = next_race_week - week
        description += f"- Next race: {next_race_name} (a **{next_race_cat}**-event) within {weeks_to_go} weeks.\n\n "

    if week not in description_added:
        description_added[week] = False
        
    if description.strip() and not description_added[week]:
        create_update_or_delete_note_event(start_date, description, note_color, events)
        description_added[week] = True
