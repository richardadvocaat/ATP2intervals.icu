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

def get_note_color(period):
    base_period = period.split()[0]
    color_mapping = {
        "Base": "yellow",
        "Peak": "orange",
        "Race": "red",
        "Transition": "green",
        "Preparation": "blue",
        "Recovery": "purple",
        "Rest": "cyan",
        "Build": "blue"
    }
    return color_mapping.get(base_period, "black")

def parse_atp_date(date_str):
    for fmt in ("%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(date_str), fmt)
        except ValueError:
            continue
    raise ValueError(f"Date '{date_str}' is not in a recognized format.")

def read_ATP_period(ATP_file_path, sheet_name="ATP_Conditions"):
    df_cond = pd.read_excel(ATP_file_path, sheet_name=sheet_name, usecols="B:C")
    cond_dict = dict(zip(df_cond.iloc[:, 0], df_cond.iloc[:, 1]))
    start_str = cond_dict.get("Start_ATP")
    end_str = cond_dict.get("End_ATP")
    start_date = parse_atp_date(start_str)
    end_date = parse_atp_date(end_str)
    oldest_date = start_date.strftime("%Y-%m-%dT00:00:00")
    newest_date = end_date.strftime("%Y-%m-%dT00:00:00")
    return oldest_date, newest_date

def get_last_day_of_week(date):
    return date + timedelta(days=(6 - date.weekday()))

def get_period_end_date(df, start_index):
    period = df.at[start_index, 'period']
    for i in range(start_index, len(df)):
        if df.at[i, 'period'] != period:
            return get_last_day_of_week(df.at[i-1, 'start_date_local'])
    return get_last_day_of_week(df.at[len(df)-1, 'start_date_local'])

def handle_period_name(period):
    period = period.strip()
    if period == "Trans":
        return "Transition"
    elif period == "Prep":
        return "Preparation"
    elif period and not period[-1].isdigit():
        return period.strip()
    return period

def get_first_a_event(df, note_event_date):
    note_date = datetime.strptime(note_event_date, "%Y-%m-%dT00:00:00")
    for index, row in df.iterrows():
        event_date = pd.to_datetime(row.get('start_date_local'))
        if event_date > note_date and str(row.get('cat', '')).upper() == 'A' and row.get('race', '').strip():
            return row.get('race', '').strip()
    return None

def populate_race_description(description, first_a_event):
    if first_a_event:
        description = f"This (part) of the plan aims for **{first_a_event}**.\n\n" + description
    return description

def create_description(period, start_date, end_date, first_a_event):
    period = handle_period_name(period)
    description = f"You are in the **{period}-period** (from {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')}).\n\n"
    description = populate_race_description(description, first_a_event)
    description += note_underline_PERIOD
    return description

def get_existing_period_notes(athlete_id, username, api_key, oldest_date, newest_date, note_name_PREFIX):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": "NOTE"}
    response = call_with_retries(requests.get, url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        events = response.json()
        logging.info(f"Fetched existing period NOTE events for athlete {athlete_id}")
        return {ev['name']: ev for ev in events if ev['name'].startswith(note_name_PREFIX)}
    logging.error(f"Failed to fetch existing period NOTE events: {response.status_code}")
    return {}

def update_note_event(event_id, start_date, end_date, description, period, athlete_id, username, api_key):
    url_base_local = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
    url_put = f"{url_base_local}/events/{event_id}"
    API_headers_local = {"Content-Type": "application/json"}
    color = get_note_color(period)
    put_data = {
        "category": "NOTE",
        "start_date_local": start_date.strftime("%Y-%m-%dT00:00:00"),
        "end_date_local": (end_date + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"),
        "name": f"{note_name_PERIOD} {handle_period_name(period)}",
        "description": description,
        "color": color
    }
    response_put = call_with_retries(requests.put, url_put, headers=API_headers_local, json=put_data, auth=HTTPBasicAuth(username, api_key))
    if response_put.status_code == 200:
        logging.info(f"Updated period NOTE event: {put_data['name']}")
    else:
        logging.error(f"Error updating period NOTE event: {put_data['name']}, code={response_put.status_code}")

def create_note_event(start_date, end_date, description, period, athlete_id, username, api_key):
    url_base_local = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
    url_post = f"{url_base_local}/events"
    API_headers_local = {"Content-Type": "application/json"}
    color = get_note_color(period)
    post_data = {
        "category": "NOTE",
        "start_date_local": start_date.strftime("%Y-%m-%dT00:00:00"),
        "end_date_local": (end_date + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"),
        "name": f"{note_name_PERIOD} {handle_period_name(period)}",
        "description": description,
        "color": color
    }
    response_post = call_with_retries(requests.post, url_post, headers=API_headers_local, json=post_data, auth=HTTPBasicAuth(username, api_key))
    if response_post.status_code == 200:
        logging.info(f"Created period NOTE event: {post_data['name']}")
    else:
        logging.error(f"Error creating period NOTE event: {post_data['name']}, code={response_post.status_code}")

def delete_note_event(event_id, athlete_id, username, api_key):
    url_del = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
    response_del = call_with_retries(requests.delete, url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
    if response_del.status_code == 200:
        logging.info(f"Deleted period NOTE event ID={event_id}")
    else:
        logging.error(f"Error deleting period NOTE event ID={event_id}: {response_del.status_code}")

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name, engine='openpyxl')
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], format='%d-%b')
    df['period'] = df['period'].astype(str)
    df['start_date_local'] = df['start_date_local'].astype('datetime64[ns]')
    df.fillna('', inplace=True)

    oldest_date, newest_date = read_ATP_period(ATP_file_path)
    oldest = pd.to_datetime(oldest_date)
    newest = pd.to_datetime(newest_date)
    df = df[(df['start_date_local'] >= oldest) & (df['start_date_local'] <= newest)]

    oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
    newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

    # Fetch existing period notes
    existing_notes = get_existing_period_notes(athlete_id, username, api_key, oldest_date, newest_date, note_name_PERIOD)

    # Determine desired period notes
    desired_notes = {}
    for i in range(len(df)):
        start_date = df.at[i, 'start_date_local']
        period = df.at[i, 'period']
        if i == 0 or df.at[i-1, 'period'] != period:
            end_date = get_period_end_date(df, i)
            first_a_event = get_first_a_event(df, start_date.strftime("%Y-%m-%dT00:00:00"))
            name = f"{note_name_PERIOD} {handle_period_name(period)}"
            description = create_description(period, start_date, end_date, first_a_event)
            desired_notes[name] = {
                "start_date": start_date,
                "end_date": end_date,
                "description": description,
                "period": period
            }

    # Sync notes: update/create/delete as needed
    for name, note in desired_notes.items():
        ex_note = existing_notes.get(name)
        if ex_note:
            # Only update if content is different
            if ex_note.get("description") != note["description"]:
                logging.info(f"Updating period NOTE event: {name}")
                update_note_event(ex_note['id'], note["start_date"], note["end_date"], note["description"], note["period"], athlete_id, username, api_key)
            else:
                logging.info(f"No update needed for period NOTE event: {name}")
        else:
            logging.info(f"Creating period NOTE event: {name}")
            create_note_event(note["start_date"], note["end_date"], note["description"], note["period"], athlete_id, username, api_key)

    # Delete obsolete notes
    for name, ex_note in existing_notes.items():
        if name not in desired_notes:
            logging.info(f"Deleting obsolete period NOTE event: {name}")
            delete_note_event(ex_note['id'], athlete_id, username, api_key)

if __name__ == "__main__":
    main()
