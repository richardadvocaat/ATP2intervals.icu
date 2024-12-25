import logging
import os
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module  # Rename the time module to avoid conflict

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(level)s - %(message)s')

# Function to format activity name to camel case
def format_activity_name(activity):
    """
    Converts an activity string with underscores to camel case.

    Args:
        activity (str): Activity name with underscores.

    Returns:
        str: Activity name in camel case.
    """
    return ''.join(word.capitalize() for word in activity.split('_'))

# Function to read user data from an Excel file
def read_user_data(excel_file_path, sheet_name="userdata"):
    """
    Reads user data from an Excel file.

    Args:
        excel_file_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet containing user data.

    Returns:
        dict: User data as a dictionary.
    """
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
sheet_name = os.getenv('SHEET_NAME', "ATP")
whattodowithrest = "**Stay in bed or on the beach! :-)**"
note_color = "red"
note_name = "Weekly Summary"
parse_delay = .01

# API endpoints
url_base = "https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
HEADERS = {"Content-Type": "application/json"}

# Function to get athlete's name
def get_athlete_name(athlete_id, username, api_key):
    """
    Fetches the athlete's name using their profile data from the API.

    Args:
        athlete_id (str): Athlete ID.
        username (str): API username.
        api_key (str): API key.

    Returns:
        str: Athlete's first name.
    """
    response = requests.get(url_profile, auth=HTTPBasicAuth(username, api_key), headers=HEADERS)
    logging.info(f"Response Status Code: {response.status_code}")
    logging.info(f"Response Headers: {response.headers}")
    logging.info(f"Response Text: {response.text}")
    if response.status_code == 200:
        profile = response.json()
        full_name = profile.get('athlete', {}).get('name', 'Athlete without name')
        first_name = full_name.split()[0] if full_name else 'Athlete'
        return first_name
    else:
        logging.error(f"Error fetching athlete profile: {response.status_code}")
        try:
            logging.error(f"Response JSON: {response.json()}")
        except ValueError:
            logging.error("Response content is not in JSON format")
        raise Exception(f"Error fetching athlete profile: {response.status_code}")

athlete_name = get_athlete_name(athlete_id, username, api_key)
print(f"Athlete First Name: {athlete_name}")

logging.info(f"Using athlete first name: {athlete_name} for further processing.")

# Conversion factors
CONVERSION_FACTORS = {
    "metric": 1000,
    "imperial": 1609.344
}

# Function to delete events within a specified date range and category
def delete_events(athlete_id, username, api_key, oldest_date, newest_date, category, name=None):
    """
    Deletes events within a specified date range and category, optionally filtered by name.

    Args:
        athlete_id (str): Athlete ID.
        username (str): API username.
        api_key (str): API key.
        oldest_date (str): Oldest date in ISO format.
        newest_date (str): Newest date in ISO format.
        category (str): Event category.
        name (str, optional): Event name to filter by.
    """
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": category}
    response_get = requests.get(url_get, headers=HEADERS, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    for event in events:
        if name and event['name'] != name:
            continue
        event_id = event['id']
        url_del = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
        response_del = requests.delete(url_del, headers=HEADERS, auth=HTTPBasicAuth(username, api_key))
        if response_del.status_code == 200:
            logging.info(f"Deleted {category.lower()} event ID={event_id}")
        else:
            logging.error(f"Error deleting {category.lower()} event ID={event_id}: {response_del.status_code}")
        time_module.sleep(parse_delay)  # Add delay between each delete event

# Function to create, update, or delete a target event based on the provided data
def create_update_or_delete_target_event(start_date, load_target, time_target, distance_target, activity_type, events, athlete_id, username, api_key):
    """
    Creates, updates, or deletes a target event based on the provided data.

    Args:
        start_date (str): Start date in ISO format.
        load_target (int): Load target.
        time_target (int): Time target.
        distance_target (int): Distance target.
        activity_type (str): Activity type.
        events (list): List of existing events.
        athlete_id (str): Athlete ID.
        username (str): API username.
        api_key (str): API key.
    """
    if load_target is None or load_target == 0:
        logging.info(f"Skipping {activity_type} event on {start_date} due to None or 0 load target.")
        return

    load_target = load_target or 0
    time_target = time_target or 0
    distance_target = distance_target or 0

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

        if server_load_target != load_target or server_time_target != time_target or server_distance_target != distance_target:
            url_put = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
            put_data = {
                "load_target": load_target,
                "time_target": time_target,
                "distance_target": distance_target
            }
            logging.info(f"Updating event: ID={event_id}, Data={put_data}")
            response_put = requests.put(url_put, headers=HEADERS, json=put_data, auth=HTTPBasicAuth(username, api_key))
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
            response_post = requests.post(url_post, headers=HEADERS, json=post_data, auth=HTTPBasicAuth(username, api_key))
            if response_post.status_code == 200:
                logging.info(f"New event created for {activity_type} on {start_date}!")
            else:
                logging.error(f"Error creating event for {activity_type} on {start_date}: {response_post.status_code}")
            time_module.sleep(parse_delay)  # Add delay between each add event

# Function to create, update, or delete a note event based on the provided data
def create_update_or_delete_note_event(start_date, description, color, events, athlete_id, username, api_key):
    """
    Creates, updates, or deletes a note event based on the provided data.

    Args:
        start_date (str): Start date in ISO format.
        description (str): Event description.
        color (str): Note color.
        events (list): List of existing events.
        athlete_id (str): Athlete ID.
        username (str): API username.
        api_key (str): API key.
    """
    end_date = start_date

    description = populate_description(description)  # Use the new function to populate the description

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
        "color": color
    }

    logging.info(f"New event: Data={post_data}")
    url_post = f"{url_base}/events".format(athlete_id=athlete_id)
    response_post = requests.post(url_post, headers=HEADERS, json=post_data, auth=HTTPBasicAuth(username, api_key))
    if response_post.status_code == 200:
        logging.info(f"New event created on {start_date}!")
    else:
        logging.error(f"Error creating event on {start_date}: {response_post.status_code}")
        time_module.sleep(parse_delay)  # Add delay between requests

# Function to format focus items into a readable list
def format_focus_items_notes(focus_items_notes):
    """
    Formats focus items into a readable list.

    Args:
        focus_items_notes (list): List of focus items.

    Returns:
        str: Formatted focus items.
    """
    if len(focus_items_notes) > 1:
        return ', '.join(focus_items_notes[:-1]) + ' and ' + focus_items_notes[-1]
    return ''.join(focus_items_notes)

# Function to populate the description
def populate_description(description):
    if not description:
        description = "Nothing to mention this week."
        
    description = f"Hi **{athlete_name}**, here is your weekly summary:\n\n" + description
    return description

# Function to format focus items into a readable list
def format_focus_items_notes(focus_items_notes):
    """
    Formats focus items into a readable list.

    Args:
        focus_items_notes (list): List of focus items.

    Returns:
        str: Formatted focus items.
    """
    if len(focus_items_notes) > 1:
        return ', '.join(focus_items_notes[:-1]) + ' and ' + focus_items_notes[-1]
    return ''.join(focus_items_notes)

# Function to handle period description
def add_period_description(row, description):
    period = row['period'] if not pd.isna(row['period']) else ""
    if period:
        description += f"- You are in the **{period}** period of your trainingplan.\n\n"
        if period == "Rest":
            description += f"- {whattodowithrest}\n\n"
    return description

# Function to handle test description
def add_test_description(row, description):
    test = row['test'] if not pd.isna(row['test']) else ""
    if test:
        description += f"- Do the following test(s) this week: **{test}**.\n\n"
    return description

# Function to handle focus description
def add_focus_description(row, description):
    focus_columns = [
        'Aerobic Endurance', 'Muscular force', 'Speed Skills',
        'Muscular Endurance', 'Anaerobic Endurance', 'Sprint Power'
    ]
    additional_focus = [col for col in focus_columns if str(row.get(col, '')).lower() == 'x']
    if additional_focus:
        formatted_focus = format_focus_items_notes(additional_focus)
        description += f"- Focus on **{formatted_focus}**.\n\n"
    elif description.strip():
        description += "- You don't have to focus on specific workouts this week.\n\n"
    return description

# Function to handle race focus description
def add_race_focus_description(row, description):
    race_cat = str(row.get('cat', '')).upper()
    race_name = row.get('race', '').strip()
    if race_cat == 'A' and race_name:
        description += f"- Use the **{race_name}** as an {race_cat}-event to primarily focus this week on this race.\n\n"
    elif race_cat == 'B' and race_name:
        description += f"- Use the **{race_name}** to learn and improve skills.\n\n"
    elif race_cat == 'C' and race_name:
        description += f"- Use the **{race_name}** as a hard effort training or just having fun!\n\n"
    return description

# Function to handle next race description
def add_next_race_description(index, df, week, description):
    next_race = None
    for i in range(index + 1, len(df)):
        next_race_name = df.at[i, 'race']
        if next_race_name and next_race_name not in ['-', '0', 'None']:
            next_race = df.iloc[i]
            break
    if next_race is not None:
        next_race_date = pd.to_datetime(next_race['race_date']).strftime("%Y-%m-%dT00:00:00")
        next_race_month = pd.to_datetime(next_race['race_date']).strftime("%B")
        next_race_week = pd.to_datetime(next_race['race_date']).isocalendar()[1]
        next_race_day = pd.to_datetime(next_race['race_date']).strftime("%A")
        next_race_dayofmonth = pd.to_datetime(next_race['race_date']).day
        next_race_name = next_race.get('race', '').strip()
        next_race_cat = str(next_race.get('cat', '')).upper()
        weeks_to_go = next_race_week - week
        if weeks_to_go == 1:
            description += f"- Upcoming race: {next_race_name} (a **{next_race_cat}**-event) next week on {next_race_day} {next_race_dayofmonth} {next_race_month}.\n\n "
        if weeks_to_go > 1:
            description += f"- Upcoming race: {next_race_name} (a **{next_race_cat}**-event) within **{weeks_to_go}** weeks on {next_race_day} {next_race_dayofmonth} {next_race_month}.\n\n "    
    return description

# Main function to execute the script logic
def main():
    """
    Main function to execute the script logic.
    """
    user_data = read_user_data(r'C:\Temp\USERDATA.xlsx')
    excel_file_path = user_data.get('EXCEL_FILE_PATH', r"C:\TEMP\ATP.xlsx")
    api_key = user_data.get('API_KEY', "yourapikey")
    username = user_data.get('USERNAME', "API_KEY")
    athlete_id = user_data.get('ATHLETE_ID', "athleteid")
    sheet_name = os.getenv('SHEET_NAME', "ATP")

    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    df.fillna(0, inplace=True)

    oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
    newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

    delete_events(athlete_id, username, api_key, oldest_date, newest_date, "TARGET")
    delete_events(athlete_id, username, api_key, oldest_date, newest_date, "NOTE", note_name)

    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET,NOTE", "resolve": "false"}
    response_get = requests.get(url_get, headers=HEADERS, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

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
                    create_update_or_delete_target_event(start_date, load, time, distance, activity, events, athlete_id, username, api_key)
                elif any([event['type'] == activity for event in events if event['start_date_local'] == start_date]):
                    create_update_or_delete_target_event(start_date, load, time, distance, activity, events, athlete_id, username, api_key)

        if all(row[col] == 0 for col in df.columns if col.endswith('_load')) and all(row[col] == 0 for col in df.columns if col.endswith('_time')) and all(row[col] == 0 for col in df.columns if col.endswith('_distance')):
            for col in df.columns:
                if col.endswith('_load'):
                    activity = format_activity_name(col.split('_load')[0])
                    create_update_or_delete_target_event(start_date, 0, 0, 0, activity, events, athlete_id, username, api_key)
            time_module.sleep(parse_delay)  # Add delay between each loop iteration for target events

    description_added = {}
    for index, row in df.iterrows():
        start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
        week = row['start_date_local'].isocalendar()[1]
        description = ""
        description = add_period_description(row, description)
        description = add_test_description(row, description)
        description = add_focus_description(row, description)
        race_focus_description = add_race_focus_description(row, description)
        if race_focus_description == description:  # If race focus description didn't change
            description = add_next_race_description(index, df, week, description)
        else:
            description = race_focus_description

        if week not in description_added:
            description_added[week] = False

        if description.strip() and not description_added[week]:
            create_update_or_delete_note_event(start_date, description, note_color, events, athlete_id, username, api_key)
            description_added[week] = True
        time_module.sleep(parse_delay)  # Add delay between each loop iteration for note events

if __name__ == "__main__":
    main()  
