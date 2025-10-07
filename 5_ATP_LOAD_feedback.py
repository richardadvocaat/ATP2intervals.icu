import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta

# --- Configuration ---
Athlete_TLA = "TLA" # Three letter Acronym of athlete.
ATP_year = "YYYY"
ATP_sheet_name = "ATP_Data"
ATP_sheet_Conditions = "ATP_Conditions"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\ATP2intervals_{Athlete_TLA}_{ATP_year}.xlsm"
parse_delay = .01
note_FEEDBACK_name_template = "Weekly feedback about your training in week {last_week}"
#NOTES_underline = "\n---\n*made with the 5_ATP_LOAD_feedback.py script / From coach Joe*"  # fill "" if you want to leave it blank.
NOTES_underline = "-"
compliance_treshold = 0.3

# --- Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def format_activity_name(activity):
    return ''.join(word.capitalize() for word in activity.split('_'))

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
    return start_date, end_date, oldest_date, newest_date

def get_previous_week(year, week):
    if week == 1:
        return year - 1, 52
    else:
        return year, week - 1

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")
unit_preference = user_data.get('DISTANCE_SYSTEM', "metric")
note_ATP_color = user_data.get('NOTE_ATP_COLOR', "red")
note_FEEDBACK_color = user_data.get('NOTE_FEEDBACK_COLOR', "blue")

url_base = "https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
url_activities = f"https://intervals.icu/api/v1/athlete/{athlete_id}/activities"
API_headers = {"Content-Type": "application/json"}

# --- Date Handling Based on ATP Period ---
start_atp_date, end_atp_date, oldest_date_str, newest_date_str = read_ATP_period(ATP_file_path, sheet_name=ATP_sheet_Conditions)
oldest_date = start_atp_date
newest_date = end_atp_date

def get_athlete_name(athlete_id, username, api_key):
    response = requests.get(url_profile, auth=HTTPBasicAuth(username, api_key), headers=API_headers)
    logging.info(f"Response Status Code: {response.status_code}")
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
logging.info(f"Using athlete first name: {athlete_name} for further processing.")

def get_wellness_data(athlete_id, username, api_key, oldest_date, newest_date):
    url_wellness = f"https://intervals.icu/api/v1/athlete/{athlete_id}/wellness"
    response = requests.get(url_wellness, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        data = response.json()
        # Filter to include only date, ctlLoad, and atlLoad within the date range
        filtered_data = [
            { "id": entry["id"], "ctlLoad": entry.get("ctlLoad", 0), "atlLoad": entry.get("atlLoad", 0) }
            for entry in data
            if oldest_date <= datetime.strptime(entry["id"], "%Y-%m-%d") <= newest_date
        ]
        return filtered_data
    else:
        logging.error(f"Error fetching wellness data: {response.status_code}")
        return []

def calculate_weekly_loads(wellness_data):
    weekly_loads = {}
    for entry in wellness_data:
        if 'id' not in entry:
            continue
        date = datetime.strptime(entry['id'], "%Y-%m-%d")
        week = date.isocalendar()[1]
        year = date.isocalendar()[0]
        year_week = f"{year}-{week}"
        if year_week not in weekly_loads:
            weekly_loads[year_week] = {'ctlLoad': 0, 'atlLoad': 0}
        weekly_loads[year_week]['ctlLoad'] += round(entry.get('ctlLoad', 0))
        weekly_loads[year_week]['atlLoad'] += round(entry.get('atlLoad', 0))
        logging.debug(f"Year-Week {year_week}: ctlLoad={weekly_loads[year_week]['ctlLoad']}, atlLoad={weekly_loads[year_week]['atlLoad']}")
    return weekly_loads

def find_existing_note_event(events, feedback_note_name):
    for event in events:
        if event.get('name') == feedback_note_name and event.get('category') == 'NOTE':
            return event
    return None

def create_update_or_delete_note_event(start_date, description, color, events, athlete_id, username, api_key, last_week, existing_note_event=None):
    end_date = start_date
    description = populate_description(description)
    post_data = {
        "category": "NOTE",
        "start_date_local": start_date,
        "end_date_local": end_date,
        "name": note_FEEDBACK_name_template.format(last_week=last_week),
        "description": description,
        "not_on_fitness_chart": "true",
        "show_as_note": "false",
        "show_on_ctl_line": "false",
        "athlete_cannot_edit": "false",
        "color": color,
        "for_week": "true"
    }
    if existing_note_event:
        # Update existing note
        url_update = f"{url_base}/events/{existing_note_event['id']}".format(athlete_id=athlete_id)
        put_data = {
            "description": description,
            "color": color
        }
        response_put = requests.put(url_update, headers=API_headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
        if response_put.status_code == 200:
            logging.info(f"Updated event NOTE for {start_date}!")
        else:
            logging.error(f"Error updating event NOTE on {start_date}: {response_put.status_code}")
        time_module.sleep(parse_delay)
    else:
        # Create new note
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
    description = f"Hi {athlete_name}, here is your weekly feedback on your training:\n\n" + description
    description += NOTES_underline
    return description

def add_load_check_description(row, previous_week_loads, previous_week_sheet_load, description):
    ctl_load = round(previous_week_loads['ctlLoad'])
    atl_load = round(previous_week_loads['atlLoad'])
    delta_ctl = ctl_load - previous_week_sheet_load
    delta_atl = atl_load - previous_week_sheet_load
    feedback = "You were compliant to the ATP! Keep doing this!"
    if previous_week_sheet_load == 0 and ctl_load == 0 and atl_load == 0:
        feedback = 'Nothing to check now.'
    elif ctl_load == 0 and atl_load == 0:
        feedback = "Nothing done? Let's talk about this."
    elif previous_week_sheet_load == 0:
        feedback = "There was nothing to do this week. :-)"
    elif delta_ctl == 0 or delta_atl == 0:
        feedback = "You were perfectly compliant to the ATP!"
    elif delta_ctl > compliance_treshold * previous_week_sheet_load or delta_atl > compliance_treshold * previous_week_sheet_load:
        feedback = "You did too much. No problem, but be aware of overreaching."
    elif delta_ctl < -compliance_treshold * previous_week_sheet_load or delta_atl < -compliance_treshold * previous_week_sheet_load:
        feedback = "You did too little. No problem, but don't make a habit of it."
    description += f"\n\nYour total trainingload for the last week was: {ctl_load}. Compared to the planned load: {previous_week_sheet_load}. Feedback: {feedback}"
    return description

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], errors='coerce')
    df = df.dropna(subset=['start_date_local'])
    df = df[(df['start_date_local'] >= oldest_date) & (df['start_date_local'] <= newest_date)]
    df['year_week'] = df['start_date_local'].apply(lambda x: f"{x.isocalendar()[0]}-{x.isocalendar()[1]}")
    start_week = start_atp_date.isocalendar()[1]
    start_year = start_atp_date.isocalendar()[0]
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {
        "oldest": oldest_date.strftime("%Y-%m-%dT00:00:00"),
        "newest": newest_date.strftime("%Y-%m-%dT00:00:00")
    }
    response = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response.json() if response.status_code == 200 else []
    today = datetime.today().date()
    for index, row in df.iterrows():
        start_date = row['start_date_local'].date()
        if start_date > today:
            continue
        start_date_str = start_date.strftime("%Y-%m-%dT00:00:00")
        week = row['start_date_local'].isocalendar()[1]
        year = row['start_date_local'].isocalendar()[0]
        previous_year, previous_week = get_previous_week(year, week)
        previous_year_week = f"{previous_year}-{previous_week}"
        previous_week_sheet_load = df[df['year_week'] == previous_year_week]['Total_load_target'].sum()
        wellness_data = get_wellness_data(athlete_id, username, api_key, oldest_date, newest_date)
        weekly_loads = calculate_weekly_loads(wellness_data)
        previous_week_loads = weekly_loads.get(previous_year_week, {'ctlLoad': 0, 'atlLoad': 0})
        feedback_note_name = note_FEEDBACK_name_template.format(last_week=previous_week)
        existing_note_event = find_existing_note_event(events, feedback_note_name)
        if year == start_year and week == start_week:
            current_description = "No feedback for the first week of the ATP"
        else:
            current_description = add_load_check_description(row, previous_week_loads, previous_week_sheet_load, "")
        new_full_description = populate_description(current_description)
        if existing_note_event:
            if existing_note_event.get('description', '') != new_full_description:
                create_update_or_delete_note_event(
                    start_date_str, current_description, note_FEEDBACK_color, events,
                    athlete_id, username, api_key, previous_week, existing_note_event=existing_note_event
                )
            else:
                logging.info(f"No update needed for feedback NOTE {feedback_note_name}.")
            continue
        else:
            if new_full_description.strip():
                create_update_or_delete_note_event(
                    start_date_str, current_description, note_FEEDBACK_color, events,
                    athlete_id, username, api_key, previous_week, existing_note_event=None
                )
        time_module.sleep(parse_delay)

if __name__ == "__main__":
    main()
