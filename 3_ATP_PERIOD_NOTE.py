from ATP_common_config import *
import time
import random
import re

# --- API Rate Limiting and Retry Logic (from 1_ATP_LOAD.py) ---
MAX_RETRIES = 4
INITIAL_BACKOFF = 0.5  # seconds
MAX_BACKOFF = 8.0      # seconds
RATE_LIMIT_DELAY = 0.25  # seconds

def call_with_retries(request_func, *args, **kwargs):
    """Call an API function with retries and exponential backoff."""
    delay = INITIAL_BACKOFF
    for attempt in range(MAX_RETRIES):
        response = request_func(*args, **kwargs)
        if response.status_code in (200, 201, 204):
            time.sleep(RATE_LIMIT_DELAY)
            return response
        elif response.status_code in (429, 500, 502, 503, 504):  # Retryable errors
            logging.warning(f"API call failed with {response.status_code}, retry #{attempt + 1} after {delay}s.")
            time.sleep(delay + random.uniform(0, 0.25))
            delay = min(MAX_BACKOFF, delay * 2)
        else:
            logging.error(f"API call failed with {response.status_code}: {getattr(response, 'text', '')}")
            break
    return response  # Return last response for error handling

def get_note_color(period):
    """
    Map period names (including numbered variants) to colors.

    Requested mapping:
      - Base 1: lightblue
      - Base 2: blue
      - Base 3: purple
      - Build 1: orange
      - Build 2: red
      - Peak: magenta
      - Race: yellow
      - Transition or Rest: lightgreen

    If an unnumbered 'Base' or 'Build' appears, a sensible default is chosen:
      - 'Base' -> lightblue
      - 'Build' -> orange

    Unknown periods fall back to 'black'.
    """
    if not period:
        return "black"

    p = str(period).strip()
    p_lower = p.lower()

    # Base N (1-3)
    m = re.match(r"^base\s*([1-3])$", p_lower)
    if m:
        return {"1": "lightblue", "2": "blue", "3": "purple"}[m.group(1)]

    # Build N (1-2)
    m = re.match(r"^build\s*([1-2])$", p_lower)
    if m:
        return {"1": "orange", "2": "red"}[m.group(1)]

    # Exact or prefix matches
    if p_lower.startswith("peak"):
        return "magenta"
    if p_lower.startswith("race"):
        return "yellow"
    if p_lower.startswith("transition") or p_lower.startswith("rest"):
        return "lightgreen"

    # Fallback sensible defaults for unnumbered variants
    if p_lower.startswith("base"):
        return "lightblue"
    if p_lower.startswith("build"):
        return "orange"

    return "black"

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

def delete_events(athlete_id, username, api_key, oldest_date, newest_date, category, name_prefix):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": category}
    response_get = call_with_retries(
        requests.get, url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key)
    )
    events = response_get.json() if response_get.status_code == 200 else []

    for event in events:
        if name_prefix and not event['name'].startswith(name_prefix):
            continue
        event_id = event['id']
        url_del = f"{url_base}/events/{event_id}".format(athlete_id=athlete_id)
        response_del = call_with_retries(
            requests.delete, url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key)
        )
        if response_del.status_code == 200:
            logging.info(f"Deleted {category.lower()} event ID={event_id}")
        else:
            logging.error(f"Error deleting {category.lower()} event ID={event_id}: {response_del.status_code}")

def handle_period_name(period):
    """Returns the cleaned, mapped period name for use in NOTE events and logic."""
    period = period.strip()
    # Map abbreviations to full names
    abbreviation_map = {
        "Trans": "Transition",
        "Prep": "Preparation",
    }
    # If abbreviation, map
    if period in abbreviation_map:
        return abbreviation_map[period]
    return period

def create_note_event(start_date, end_date, description, period_name, athlete_id, username, api_key):
    url_base = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
    url_post = f"{url_base}/events"
    API_headers = {"Content-Type": "application/json"}

    color = get_note_color(period_name)

    post_data = {
        "category": "NOTE",
        "start_date_local": start_date.strftime("%Y-%m-%dT00:00:00"),
        "end_date_local": (end_date + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"),  # Add an extra day
        "name": f"{note_name_PERIOD} {period_name}",
        "description": description,
        "color": color
    }

    response_post = call_with_retries(
        requests.post, url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key)
    )
    if response_post.status_code == 200:
        logging.info(f"New event created from {start_date} to {end_date}!")
    else:
        logging.error(f"Error creating event: {response_post.status_code} - {response_post.text}")

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

def create_description(period_name, start_date, end_date, first_a_event):
    description = f"You are in the **{period_name}-period** (from {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')}).\n\n"
    description = populate_race_description(description, first_a_event)
    description += note_underline_PERIOD
    return description

def get_existing_period_notes(athlete_id, oldest_date, newest_date, username, api_key, note_name_PERIOD):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date, "newest": newest_date, "category": "NOTE"}
    response = call_with_retries(
        requests.get, url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key)
    )
    notes = response.json() if response.status_code == 200 else []
    # Only pick notes with correct prefix
    period_notes = {}
    for note in notes:
        if note['name'].startswith(note_name_PERIOD):
            key = (
                note['start_date_local'],
                note['end_date_local'],
                note['name']
            )
            period_notes[key] = note
    return period_notes

def get_desired_period_notes(df):
    desired_notes = {}
    for i in range(len(df)):
        start_date = df.at[i, 'start_date_local']
        period = df.at[i, 'period']
        # Use cleaned full period name
        period_name = handle_period_name(period)
        if i == 0 or df.at[i-1, 'period'] != period:
            end_date = get_period_end_date(df, i)
            first_a_event = get_first_a_event(df, start_date.strftime("%Y-%m-%dT00:00:00"))
            description = create_description(period_name, start_date, end_date, first_a_event)
            name = f"{note_name_PERIOD} {period_name}"
            color = get_note_color(period_name)
            key = (
                start_date.strftime("%Y-%m-%dT00:00:00"),
                (end_date + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00"),
                name
            )
            desired_notes[key] = {
                "category": "NOTE",
                "start_date_local": key[0],
                "end_date_local": key[1],
                "name": name,
                "description": description,
                "color": color,
                "period_name": period_name  # For later use
            }
    return desired_notes

def main():
    oldest_date, newest_date = read_ATP_period(ATP_file_path)
    oldest = pd.to_datetime(oldest_date)
    newest = pd.to_datetime(newest_date)
    now = datetime.now()

    # Only prompt if ATP period includes today or past
    if oldest < now:
        answer = input("Do you want to delete notes in the past? (yes/no): ").strip().lower()
        overwrite_past = answer == "yes"
    else:
        overwrite_past = True  # ATP is completely in the future, so no filter needed

    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name, engine='openpyxl')
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], format='%d-%b', errors='coerce')
    df = df.dropna(subset=['start_date_local'])

    # Strictly limit data to ATP period
    df = df[(df['start_date_local'] >= oldest) & (df['start_date_local'] <= newest)]

    # If not overwriting past, only keep future notes
    if not overwrite_past:
        df = df[df['start_date_local'] >= now]

    # Define date range for NOTE events syncing
    if df.empty:
        logging.info("No notes to process for the selected ATP period.")
        return

    oldest_date = df['start_date_local'].min().strftime("%Y-%m-%dT00:00:00")
    newest_date = df['start_date_local'].max().strftime("%Y-%m-%dT00:00:00")

    # Build the desired notes dictionary
    desired_notes = get_desired_period_notes(df)
    # Read all existing notes for this period and prefix
    existing_notes = get_existing_period_notes(athlete_id, oldest_date, newest_date, username, api_key, note_name_PERIOD)

    # 1. Update existing notes if different, or create new notes
    for key, desired_note in desired_notes.items():
        existing_note = existing_notes.get(key)
        if existing_note:
            # Compare fields: description, color
            if (
                existing_note.get("description", "") != desired_note["description"]
                or existing_note.get("color", "") != desired_note["color"]
            ):
                url_put = f"{url_base}/events/{existing_note['id']}".format(athlete_id=athlete_id)
                put_data = {
                    "description": desired_note["description"],
                    "color": desired_note["color"]
                }
                response_put = call_with_retries(
                    requests.put, url_put, headers=API_headers, json=put_data, auth=HTTPBasicAuth(username, api_key)
                )
                if response_put.status_code == 200:
                    logging.info(f"Updated NOTE {desired_note['name']}")
                else:
                    logging.error(f"Failed to update NOTE {desired_note['name']}: {response_put.status_code}")
            else:
                logging.info(f"NOTE {desired_note['name']} is unchanged; no update needed.")
        else:
            # Create new NOTE
            create_note_event(
                pd.to_datetime(desired_note["start_date_local"]),
                pd.to_datetime(desired_note["end_date_local"]) - timedelta(days=1),
                desired_note["description"],
                desired_note["period_name"],  # Pass full cleaned period name
                athlete_id, username, api_key
            )

    # 2. Delete notes that are no longer needed
    for key, existing_note in existing_notes.items():
        if key not in desired_notes:
            url_del = f"{url_base}/events/{existing_note['id']}".format(athlete_id=athlete_id)
            response_del = call_with_retries(
                requests.delete, url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key)
            )
            if response_del.status_code == 200:
                logging.info(f"Deleted NOTE {existing_note['name']}")
            else:
                logging.error(f"Failed to delete NOTE {existing_note['name']}: {response_del.status_code}")

if __name__ == "__main__":
    main()
