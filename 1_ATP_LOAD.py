from ATP_common_config import *
import time
import random
from functools import wraps

# --- API Rate Limiting and Retry Logic ---
MAX_RETRIES = 4
INITIAL_BACKOFF = 0.5  # seconds
MAX_BACKOFF = 8.0      # seconds
RATE_LIMIT_DELAY = 0.25  # seconds, adjust as needed for API

def call_with_retries(request_func, *args, **kwargs):
    """Call an API function with retries and exponential backoff."""
    delay = INITIAL_BACKOFF
    for attempt in range(MAX_RETRIES):
        response = request_func(*args, **kwargs)
        if response.status_code in (200, 201, 204):
            time.sleep(RATE_LIMIT_DELAY)  # Rate limiting after successful call
            return response
        elif response.status_code in (429, 500, 502, 503, 504):  # Retryable errors
            logging.warning(f"API call failed with {response.status_code}, retry #{attempt + 1} after {delay}s.")
            time.sleep(delay + random.uniform(0, 0.25))
            delay = min(MAX_BACKOFF, delay * 2)
        else:
            logging.error(f"API call failed with {response.status_code}: {getattr(response, 'text', '')}")
            break
    return response  # Return last response for error handling

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
    return oldest_date, newest_date

def prompt_overwrite_past():
    answer = input("Do you want to overwrite data in the past? (yes/no): ").strip().lower()
    return answer == "yes"

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")
unit_preference = user_data.get('DISTANCE_SYSTEM', "metric")
note_ATP_color = user_data.get('NOTE_ATP_COLOR', "red")
do_at_rest = user_data.get('Do_At_Rest', "Do nothing!")

url_base = "https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
url_activities = f"https://intervals.icu/api/v1/athlete/{athlete_id}/activities"
API_headers = {"Content-Type": "application/json"}

def clean_activity_name(col_name):
    return col_name.replace('_load_target', '').replace('_load', '')

def distance_conversion_factor(unit_preference):
    conversion_factors = {
        "metric": 1000,
        "imperial": 1609.344
    }
    return conversion_factors.get(unit_preference, 1000)

def normalize(val):
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return 0

def get_existing_events(athlete_id, oldest_date, newest_date, username, api_key):
    url_get = f"https://intervals.icu/api/v1/athlete/{athlete_id}/eventsjson"
    params = {"oldest": oldest_date, "newest": newest_date, "category": "TARGET"}
    response = call_with_retries(requests.get, url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        events = response.json()
        event_map = {
            (e['start_date_local'], e['type']): e
            for e in events
        }
        return event_map
    else:
        logging.error(f"Failed to fetch events ({response.status_code})")
        return {}

def get_desired_events(df):
    desired = {}
    dist_factor = distance_conversion_factor(unit_preference)
    columns = df.columns
    for row in df.itertuples(index=False):
        start_date = row.start_date_local.strftime("%Y-%m-%dT00:00:00")
        for col in columns:
            if col.endswith('_load_target'):
                activity = clean_activity_name(col)
                if activity in [None, "None", "Total"]:
                    continue
                load = normalize(getattr(row, col))
                time_col = f"{activity}_time_target"
                dist_col = f"{activity}_distance_target"
                time = normalize(getattr(row, time_col, 0)) * 60 if hasattr(row, time_col) else 0
                if hasattr(row, dist_col):
                    if activity.lower() in ['swim', 'openwaterswim']:
                        distance = normalize(getattr(row, dist_col))
                    else:
                        distance = normalize(getattr(row, dist_col)) * dist_factor
                else:
                    distance = 0
                key = (start_date, activity)
                desired[key] = {
                    'start_date_local': start_date,
                    'type': activity,
                    'load_target': load,
                    'time_target': time,
                    'distance_target': distance
                }
    return desired

def efficient_event_sync(df, athlete_id, username, api_key):
    if df.empty:
        logging.error("No valid dates found in 'start_date_local'.")
        return

    oldest_date, newest_date = read_ATP_period(ATP_file_path)
    existing_events = get_existing_events(athlete_id, oldest_date, newest_date, username, api_key)
    desired_events = get_desired_events(df)

    # 1. Create or Update events
    for key, new_event in desired_events.items():
        old_event = existing_events.get(key)
        if old_event:
            # Normalize all compared values for robust equality
            old_load = normalize(old_event.get('load_target', 0))
            old_time = normalize(old_event.get('time_target', 0))
            old_distance = normalize(old_event.get('distance_target', 0))
            new_load = normalize(new_event['load_target'])
            new_time = normalize(new_event['time_target'])
            new_distance = normalize(new_event['distance_target'])

            # Log comparison for debugging
            logging.debug(f"Comparing load_target: old={old_load}, new={new_load}")
            logging.debug(f"Comparing time_target: old={old_time}, new={new_time}")
            logging.debug(f"Comparing distance_target: old={old_distance}, new={new_distance}")

            if (
                old_load != new_load or
                old_time != new_time or
                old_distance != new_distance
            ):
                url_put = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{old_event['id']}"
                put_data = {
                    "load_target": new_event['load_target'],
                    "time_target": new_event['time_target'],
                    "distance_target": new_event['distance_target']
                }
                logging.info(f"Updating event {key}: {put_data}")
                response_put = call_with_retries(requests.put, url_put, headers=API_headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
                if response_put.status_code == 200:
                    logging.info(f"Updated event for {key}")
                else:
                    logging.error(f"Failed to update event for {key}: {response_put.status_code}")
            else:
                logging.info(f"No changes needed for event {key}")
        else:
            if any([new_event['load_target'] > 0, new_event['time_target'] > 0, new_event['distance_target'] > 0]):
                url_post = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events"
                post_data = {
                    "load_target": new_event['load_target'],
                    "time_target": new_event['time_target'],
                    "distance_target": new_event['distance_target'],
                    "category": "TARGET",
                    "type": new_event['type'],
                    "name": "Weekly",
                    "start_date_local": new_event['start_date_local']
                }
                logging.info(f"Creating event {key}: {post_data}")
                response_post = call_with_retries(requests.post, url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
                if response_post.status_code == 200:
                    logging.info(f"Created new event for {key}")
                else:
                    logging.error(f"Failed to create event for {key}: {response_post.status_code}")

    # 2. Delete events that are no longer needed
    for key, old_event in existing_events.items():
        if key not in desired_events:
            url_del = f"https://intervals.icu/api/v1/athlete/{athlete_id}/events/{old_event['id']}"
            logging.info(f"Deleting event {key}")
            response_del = call_with_retries(requests.delete, url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
            if response_del.status_code == 200:
                logging.info(f"Deleted event for {key}")
            else:
                logging.error(f"Failed to delete event for {key}: {response_del.status_code}")

def main():
    overwrite_past = prompt_overwrite_past()
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], errors='coerce')
    df = df.dropna(subset=['start_date_local'])

    oldest_date, newest_date = read_ATP_period(ATP_file_path)
    oldest = pd.to_datetime(oldest_date)
    newest = pd.to_datetime(newest_date)
    df = df[(df['start_date_local'] >= oldest) & (df['start_date_local'] <= newest)]

    if not overwrite_past:
        now = datetime.now()
        df = df[df['start_date_local'] >= now]

    efficient_event_sync(df, athlete_id, username, api_key)

if __name__ == "__main__":
    main()
