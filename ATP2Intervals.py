import logging
import os
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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
note_color = "red"
whattodowithrest = "**Stay in bed or on the beach! :-)**"
note_name = "Weekly Summary of your ATP"
# API endpoints
url_post = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events"
url_get = f"https://intervals.icu/api/v1/athlete/{athlete_id}/eventsjson"
url_delete = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events"
headers = {"Content-Type": "application/json"}

def create_update_or_delete_note_event(start_date, description, color, events):
    if not description.strip():
        logging.info(f"No description provided for {start_date}. Skipping note creation.")
        return
    
    end_date = start_date  # End date is the same as start date for NOTE events
    
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
        "color": note_color  # Use the dynamic color
    }

    duplicate_event = next((event for event in events if event['category'] == "NOTE" and event['name'] == post_data['name'] and event['start_date_local'] == post_data['start_date_local']), None)
    
    if duplicate_event:
        event_id = duplicate_event['id']
        if duplicate_event.get('description') != description or duplicate_event.get('color') != note_color:
            url_put = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{event_id}"
            put_data = {
                "description": description,
                "color": note_color  # Update the color if it has changed
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

# Read the Excel file and specify the sheet
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
df.fillna(0, inplace=True)

# Get the oldest and newest date from the Excel list
oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

# Retrieve existing events
params = {"oldest": oldest_date, "newest": newest_date, "category": "NOTE", "resolve": "false"}
response_get = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(username, api_key))
events = response_get.json() if response_get.status_code == 200 else []

# Track if description has been added for the week
description_added = {}

focus_columns = [
    'Aerobic Endurance', 'Muscular force', 'Speed Skills',
    'Muscular Endurance', 'Anaerobic Endurance', 'Sprint Power'
]

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
    description = f"You are in the **{period}** period of your trainingplan.\n\n" if period else ""
    if period == "Rest":
        description += f"{whattodowithrest}\n\n"
    if test:  # Add test comment if there is a value
        description += f"Do the following test(s) this week: **{test}**.\n\n"
    if focus:
        description += f"Focus this week on {focus}.\n\n"
    
    # Add focus based on specified columns
    additional_focus = [col for col in focus_columns if str(row.get(col, '')).lower() == 'x']
    if additional_focus:
        formatted_focus = format_focus_items_notes(additional_focus)
        description += f"Focus on {formatted_focus}.\n\n"
    elif description.strip():
        description += "You don't have to focus on specific workouts this week.\n\n"
    
    # Add focus for A, B, and C category races
    race_cat = str(row.get('cat', '')).upper()
    race_name = row.get('race', '')
    if race_cat == 'A' and race_name:
        description += f"Use the **{race_name}** as an {race_cat}-event to primarily focus this week on this race."
    elif race_cat == 'B' and race_name:
        description += f"Use the **{race_name}** to learn and improve skills."
    elif race_cat == 'C' and race_name:
        description += f"Use the **{race_name}** as hard effort training or just having fun!"

    if week not in description_added:
        description_added[week] = False
        
    if description.strip() and not description_added[week]:
        create_update_or_delete_note_event(start_date, description, note_color, events)
        description_added[week] = True
