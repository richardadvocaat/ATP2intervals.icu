import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(level)s - %(message)s')

def format_activity_name(activity):
    return ''.join(word.capitalize() for word in activity.split('_'))

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

ATP_sheet_name = "ATP_data"
ATP_file_path = r'C:\TEMP\Intervals_API_Tools_Office365_v1.6_ATP2intervals.xlsm'

parse_delay = .01
do_at_rest = "**Stay in bed, on the beach and focus on friends, family and your Märklin trainset.**"

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")
unit_preference = user_data.get('DISTANCE_SYSTEM', "metric")
note_color = user_data.get('NOTE_COLOR', "red")

url_base = "https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
url_activities = f"https://intervals.icu/api/v1/athlete/{athlete_id}/activities"
API_headers = {"Content-Type": "application/json"}

def get_athlete_name(athlete_id, username, api_key):
    response = requests.get(url_profile, auth=HTTPBasicAuth(username, api_key), headers=API_headers)
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

note_name = f"Your weekly training and focus summary"

def distance_conversion_factor(unit_preference):
    conversion_factors = {
        "metric": 1000,
        "imperial": 1609.344,
        "Rijnlands": 3.186
    }
    return conversion_factors.get(unit_preference, 1000)

def get_last_week_load(athlete_id, username, api_key):
    today = datetime.now()
    start_date = (today - timedelta(days=today.weekday() + 7)).strftime("%Y-%m-%d")
    end_date = (today - timedelta(days=today.weekday() + 1)).strftime("%Y-%m-%d")
    params = {"start": start_date, "end": end_date}
    response = requests.get(url_activities, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        activities = response.json()
        weekly_load = sum(activity.get('load', 0) for activity in activities)
        return weekly_load
    else:
        logging.error(f"Error fetching activities: {response.status_code}")
        return 0

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

    if activity_type in ["Ride", "Run"]:
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

def create_update_or_delete_note_event(start_date, description, color, events, athlete_id, username, api_key):
    end_date = start_date

    description = populate_description(description)

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
    response_post = requests.post(url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
    if response_post.status_code == 200:
        logging.info(f"New event created on {start_date}!")
    else:
        logging.error(f"Error creating event on {start_date}: {response_post.status_code}")
        time_module.sleep(parse_delay)

def format_focus_items_notes(focus_items_notes):
    if len(focus_items_notes) > 1:
        return ', '.join(focus_items_notes[:-1]) + ' and ' + focus_items_notes[-1]
    return ''.join(focus_items_notes)

def populate_description(description):
    if not description:
        description = "Nothing to mention this week."
        
    description = f"Hi **{athlete_name}**, here is your weekly summary:\n\n" + description
    return description

def add_period_description(row, description):
    period = row['period'] if not pd.isna(row['period']) else ""
    if period:
        description += f"- You are in the **{period}** period of your trainingplan.\n\n"
        if period == "Rest":
            description += f"- {do_at_rest}\n\n"
    return description

def add_test_description(row, description):
    test = row['test'] if not pd.isna(row['test']) else ""
    if test:
        description += f"- Do the following test(s) this week: **{test}**.\n\n"
    return description

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
            description += f"- Upcoming race: **{next_race_name}**(a **{next_race_cat}**-event) next week on {next_race_day} {next_race_dayofmonth} {next_race_month}.\n\n "
        if weeks_to_go > 1:
            description += f"- Upcoming race: **{next_race_name}** (a **{next_race_cat}**-event) within **{weeks_to_go}** weeks on {next_race_day} {next_race_dayofmonth} {next_race_month}.\n\n "    
    return description

def add_load_check_description(row, weekly_load_api, description):
    sheet_load = row['Total_Load'] if 'Total_Load' in row else 0
    delta = weekly_load_api - sheet_load

    if sheet_load == 0 and weekly_load_api == 0:
        feedback = '--.'
    elif weekly_load_api == 0:
        feedback = "Nothing done?"
    elif sheet_load == 0:
        feedback = "There was nothing to do...?"
    elif delta == 0:
        feedback = "Perfect!"
    elif delta > 0.2 * sheet_load:
        feedback = "Too much."
    elif delta < -0.2 * sheet_load:
        feedback = "Too little."
    else:
        feedback = "Good."
    if sheet_load == 0 and weekly_load_api == 0:
        description += f"\n\nFeedback: {feedback}"
    elif weekly_load_api == 0:
        description += f"\n\nYour total load for the last week was: {weekly_load_api}. Compared to the planned load: {sheet_load}. Feedback: {feedback}"
    
    return description

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)

    oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
    newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

    delete_events(athlete_id, username, api_key, oldest_date, newest_date, "TARGET")
    delete_events(athlete_id, username, api_key, oldest_date, newest_date, "NOTE", note_name)

    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET,NOTE", "resolve": "false"}
    response_get = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    weekly_load_api = get_last_week_load(athlete_id, username, api_key)

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
            time_module.sleep(parse_delay)

    note_description_added = {}
    for index, row in df.iterrows():
        start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
        week = row['start_date_local'].isocalendar()[1]
        description = ""
        description = add_period_description(row, description)
        description = add_test_description(row, description)
        description = add_focus_description(row, description)
        race_focus_description = add_race_focus_description(row, description)
        if race_focus_description == description:
            description = add_next_race_description(index, df, week, description)
        else:
            description = race_focus_description
        description = add_load_check_description(row, weekly_load_api, description)
            
        if week not in note_description_added:
            note_description_added[week] = False

        if description.strip() and not note_description_added[week]:
            create_update_or_delete_note_event(start_date, description, note_color, events, athlete_id, username, api_key)
            note_description_added[week] = True
        time_module.sleep(parse_delay)

if __name__ == "__main__":
    main()
    
