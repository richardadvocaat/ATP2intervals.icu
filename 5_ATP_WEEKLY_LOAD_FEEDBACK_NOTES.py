from ATP_common_config import *
import time
import random

# --- API Rate Limiting and Retry Logic ---
MAX_RETRIES = 4
INITIAL_BACKOFF = 0.5  # seconds
MAX_BACKOFF = 8.0      # seconds
RATE_LIMIT_DELAY = 0.25  # seconds

def call_with_retries(request_func, *args, **kwargs):
    delay = INITIAL_BACKOFF
    for attempt in range(MAX_RETRIES):
        response = request_func(*args, **kwargs)
        if response.status_code in (200, 201, 204):
            time.sleep(RATE_LIMIT_DELAY)
            return response
        elif response.status_code in (429, 500, 502, 503, 504):
            logging.warning(f"API call failed with {response.status_code}, retry #{attempt + 1} after {delay}s.")
            time.sleep(delay + random.uniform(0, 0.25))
            delay = min(MAX_BACKOFF, delay * 2)
        else:
            logging.error(f"API call failed with {response.status_code}: {getattr(response, 'text', '')}")
            break
    return response

def format_activity_name(activity):
    return ''.join(word.capitalize() for word in activity.split('_'))

def parse_atp_date(date_str):
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

# --- Date Handling Based on ATP Period ---
start_atp_date, end_atp_date, oldest_date_str, newest_date_str = read_ATP_period(ATP_file_path, sheet_name=ATP_sheet_Conditions)
oldest_date = start_atp_date
newest_date = end_atp_date

def get_athlete_name(athlete_id, username, api_key):
    response = call_with_retries(requests.get, url_profile, auth=HTTPBasicAuth(username, api_key), headers=API_headers)
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
    response = call_with_retries(requests.get, url_wellness, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        data = response.json()
        filtered_data = [
            { "id": entry["id"], "ctlLoad": entry.get("ctlLoad", 0), "atlLoad": entry.get("atlLoad", 0) }
            for entry in data
            if oldest_date <= datetime.strptime(entry["id"], "%Y-%m-%d") <= newest_date
        ]
        logging.info(f"Fetched wellness data for athlete {athlete_id}")
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
    logging.info("Calculated weekly loads from wellness data")
    return weekly_loads

def format_focus_items_notes(focus_items_notes):
    if len(focus_items_notes) > 1:
        return ', '.join(focus_items_notes[:-1]) + ' and ' + focus_items_notes[-1]
    return ''.join(focus_items_notes)

def populate_description(description):
    if not description:
        description = "Nothing to mention this week."
    description = f"Hi **{athlete_name}**, here is your weekly feedback on your training:\n\n" + description
    description += note_underline_FEEDBACK 
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
    description += f"- Your **total trainingload** for the last week was: **{ctl_load}**. Compared to the **planned trainingload**: **{previous_week_sheet_load}**.\n\n- **Feedback**: {feedback} \n\n"
    return description

def get_existing_feedback_notes(athlete_id, username, api_key, oldest_date, newest_date, note_name_template_FEEDBACK):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {
        "oldest": oldest_date.strftime("%Y-%m-%dT00:00:00"),
        "newest": newest_date.strftime("%Y-%m-%dT00:00:00"),
        "category": "NOTE"
    }
    response = call_with_retries(requests.get, url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        events = response.json()
        logging.info(f"Fetched existing feedback NOTE events for athlete {athlete_id}")
        return {ev['name']: ev for ev in events if ev.get('category') == 'NOTE' and ev['name'].startswith(note_name_template_FEEDBACK.split('{')[0])}
    logging.error(f"Failed to fetch existing feedback NOTE events: {response.status_code}")
    return {}

def update_note_event(event_id, start_date, description, color, athlete_id, username, api_key, last_week):
    url_put = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
    put_data = {
        "description": description,
        "color": color
    }
    response_put = call_with_retries(requests.put, url_put, headers=API_headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
    if response_put.status_code == 200:
        logging.info(f"Updated feedback NOTE event for week {last_week}")
    else:
        logging.error(f"Error updating feedback NOTE event for week {last_week}: {response_put.status_code}")

def create_note_event(start_date, description, color, athlete_id, username, api_key, last_week):
    end_date = start_date
    note_name = note_name_template_FEEDBACK.format(last_week=last_week)
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
        "color": color,
        "for_week": "true"
    }
    url_post = f"{url_base}/events".format(athlete_id=athlete_id)
    response_post = call_with_retries(requests.post, url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
    if response_post.status_code == 200:
        logging.info(f"Created feedback NOTE event for week {last_week}")
    else:
        logging.error(f"Error creating feedback NOTE event for week {last_week}: {response_post.status_code}")

def delete_note_event(event_id, athlete_id, username, api_key, last_week):
    url_del = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
    response_del = call_with_retries(requests.delete, url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
    if response_del.status_code == 200:
        logging.info(f"Deleted feedback NOTE event for week {last_week}")
    else:
        logging.error(f"Error deleting feedback NOTE event for week {last_week}: {response_del.status_code}")

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], errors='coerce')
    df = df.dropna(subset=['start_date_local'])
    df = df[(df['start_date_local'] >= oldest_date) & (df['start_date_local'] <= newest_date)]
    df['year_week'] = df['start_date_local'].apply(lambda x: f"{x.isocalendar()[0]}-{x.isocalendar()[1]}")
    start_week = start_atp_date.isocalendar()[1]
    start_year = start_atp_date.isocalendar()[0]
    today = datetime.today().date()

    # Batch fetch existing feedback NOTE events
    existing_notes = get_existing_feedback_notes(
        athlete_id, username, api_key,
        oldest_date,
        newest_date,
        note_name_template_FEEDBACK
    )

    # Determine desired feedback notes for each week
    desired_notes = {}
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
        feedback_note_name = note_name_template_FEEDBACK.format(last_week=previous_week)
        if year == start_year and week == start_week:
            current_description = "- No feedback for the first week of the ATP"
        else:
            current_description = add_load_check_description(row, previous_week_loads, previous_week_sheet_load, "")
        full_description = populate_description(current_description)
        desired_notes[feedback_note_name] = {
            "start_date": start_date_str,
            "description": full_description,
            "color": note_color_FEEDBACK,
            "week": previous_week
        }

    # Sync notes: update/create/delete as needed
    for note_name, note in desired_notes.items():
        ex_note = existing_notes.get(note_name)
        if ex_note:
            if ex_note.get('description', '') != note["description"]:
                logging.info(f"Updating feedback NOTE event: {note_name}")
                update_note_event(ex_note['id'], note["start_date"], note["description"], note["color"], athlete_id, username, api_key, note["week"])
            else:
                logging.info(f"No update needed for feedback NOTE event: {note_name}")
        else:
            logging.info(f"Creating feedback NOTE event: {note_name}")
            create_note_event(note["start_date"], note["description"], note["color"], athlete_id, username, api_key, note["week"])
        time.sleep(parse_delay)
    # Delete obsolete notes
    for note_name, ex_note in existing_notes.items():
        if note_name not in desired_notes:
            logging.info(f"Deleting obsolete feedback NOTE event: {note_name}")
            delete_note_event(ex_note['id'], athlete_id, username, api_key, note_name)
            time.sleep(parse_delay)

if __name__ == "__main__":
    main()
