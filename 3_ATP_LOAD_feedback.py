import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(level)s - %(message)s')
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def format_activity_name(activity):
    return ''.join(word.capitalize() for word in activity.split('_'))

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

ATP_sheet_name = "ATP_data"
ATP_file_path = r'C:\TEMP\Intervals_API_Tools_Office365_v1.6_ATP2intervals.xlsm'

parse_delay = .01
note_FEEDBACK_name_template = "Weekly update about your training in week {last_week}"

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

def distance_conversion_factor(unit_preference):
    conversion_factors = {
        "metric": 1000,
        "imperial": 1609.344,
        "Rijnlands": 3.186
    }
    return conversion_factors.get(unit_preference, 1000)

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
        
        weekly_loads[year_week]['ctlLoad'] += entry.get('ctlLoad', 0)
        weekly_loads[year_week]['atlLoad'] += entry.get('atlLoad', 0)
        
        logging.debug(f"Year-Week {year_week}: ctlLoad={weekly_loads[year_week]['ctlLoad']}, atlLoad={weekly_loads[year_week]['atlLoad']}")
    
    return weekly_loads

def get_weekly_loads(athlete_id, username, api_key, oldest_date, newest_date):
    wellness_data = get_wellness_data(athlete_id, username, api_key, oldest_date, newest_date)
    weekly_loads = calculate_weekly_loads(wellness_data)
    return weekly_loads

def get_last_week_load(athlete_id, username, api_key, note_event_date):
    wellness_data = get_wellness_data(athlete_id, username, api_key)
    weekly_loads = calculate_weekly_loads(wellness_data)

    note_date = datetime.strptime(note_event_date, "%Y-%m-%d")
    last_week_start = note_date - timedelta(days=note_date.weekday() + 7)
    last_week_end = last_week_start + timedelta(days=6)

    last_week_load = {'ctlLoad': 0, 'atlLoad': 0}
    for entry in wellness_data:
        date = datetime.strptime(entry['id'], "%Y-%m-%d")
        if last_week_start <= date <= last_week_end:
            last_week_load['ctlLoad'] += entry.get('ctlLoad', 0)
            last_week_load['atlLoad'] += entry.get('atlLoad', 0)

    return last_week_load

def delete_events(athlete_id, username, api_key, oldest_date, newest_date, category, name_prefix=None):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": category}
    response_get = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    for event in events:
        if name_prefix and not event['name'].startswith(name_prefix):
            continue
        event_id = event['id']
        url_del = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
        response_del = requests.delete(url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
        if response_del.status_code == 200:
            logging.info(f"Deleted {category.lower()} event ID={event_id}")
        else:
            logging.error(f"Error deleting {category.lower()} event ID={event_id}: {response_del.status_code}")
        time_module.sleep(parse_delay)


def create_update_or_delete_note_event(start_date, description, color, events, athlete_id, username, api_key, last_week):
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
        
    description = f"Hi **{athlete_name}**, here is your weekly feedback on your training:\n\n" + description
    return description

def get_previous_week(year, week):
    if week == 1:
        return year - 1, 52
    else:
        return year, week - 1
    
def calculate_total_load(row):
    return sum(row[col] for col in row.index if col.endswith('_load'))

def get_previous_week_sheet_load(df, previous_year, previous_week):
    previous_year_week = f"{previous_year}-{previous_week}"
    previous_week_data = df[(df['year_week'] == previous_year_week)]
    if not previous_week_data.empty:
        return calculate_total_load(previous_week_data.iloc[0])
    return 0

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

def add_load_check_description(row, previous_week_loads, previous_week_sheet_load, description):
    ctl_load = round(previous_week_loads['ctlLoad'])
    atl_load = round(previous_week_loads['atlLoad'])
    
    delta_ctl = ctl_load - previous_week_sheet_load
    delta_atl = atl_load - previous_week_sheet_load

    feedback = "Good."
    if previous_week_sheet_load == 0 and ctl_load == 0 and atl_load == 0:
        feedback = 'Nothing to check now.'
    elif ctl_load == 0 and atl_load == 0:
        feedback = "Nothing done?"
    elif previous_week_sheet_load == 0:
        feedback = "There was nothing to do...?"
    elif delta_ctl == 0 or delta_atl == 0:
        feedback = "Perfect!"
    elif delta_ctl > 0.2 * previous_week_sheet_load or delta_atl > 0.2 * previous_week_sheet_load:
        feedback = "You did too much."
    elif delta_ctl < -0.2 * previous_week_sheet_load or delta_atl < -0.2 * previous_week_sheet_load:
        feedback = "You did too little."

    description += f"\n\nYour total trainingload for the last week was: **{ctl_load}**. Compared to the planned load: **{previous_week_sheet_load}**. Feedback: **{feedback}**"
    
    return description

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)

    # Create 'year_week' column
    df['year_week'] = df['start_date_local'].apply(lambda x: f"{x.isocalendar()[0]}-{x.isocalendar()[1]}")

    oldest_date = df['start_date_local'].min()
    newest_date = df['start_date_local'].max()

    # Delete existing NOTE_EVENTS with the same note_FEEDBACK_name prefix before processing new ones
    delete_events(athlete_id, username, api_key, oldest_date.strftime("%Y-%m-%dT00:00:00"), newest_date.strftime("%Y-%m-%dT00:00:00"), "NOTE", note_FEEDBACK_name_template.split("{")[0])

    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date.strftime("%Y-%m-%dT00:00:00"), "newest": newest_date.strftime("%Y-%m-%dT00:00:00"), "category": "TARGET,NOTE", "resolve": "false"}
    response_get = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    events = response_get.json() if response_get.status_code == 200 else []

    weekly_loads = get_weekly_loads(athlete_id, username, api_key, oldest_date, newest_date)

    description_added = {}
    today = datetime.now().date()
    for index, row in df.iterrows():
        start_date = row['start_date_local'].date()
        if start_date > today:
            continue

        start_date_str = start_date.strftime("%Y-%m-%dT00:00:00")
        week = row['start_date_local'].isocalendar()[1]
        year = row['start_date_local'].isocalendar()[0]
        previous_year, previous_week = get_previous_week(year, week)
        previous_year_week = f"{previous_year}-{previous_week}"
        year_week = f"{year}-{week}"
                  
        description = ""
                  
        previous_week_loads = weekly_loads.get(previous_year_week, {'ctlLoad': 0, 'atlLoad': 0})
        previous_week_sheet_load = get_previous_week_sheet_load(df, previous_year, previous_week)  # Define it here
        description = add_load_check_description(row, previous_week_loads, previous_week_sheet_load, description)

        if week not in description_added:
            description_added[week] = False

        if description.strip() and not description_added[week]:
            create_update_or_delete_note_event(start_date_str, description, note_FEEDBACK_color, events, athlete_id, username, api_key, previous_week)
            description_added[week] = True
        time_module.sleep(parse_delay)

if __name__ == "__main__":
    main()
