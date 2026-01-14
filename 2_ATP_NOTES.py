from ATP_common_config import *
import time
import random

# --- API Rate Limiting and Retry Logic ---
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

def format_activity_name(activity):
    return ''.join(word.capitalize() for word in activity.split('_'))

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

def get_athlete_name(athlete_id, username, api_key):
    response = call_with_retries(requests.get, url_profile, auth=HTTPBasicAuth(username, api_key), headers=API_headers)
    if response.status_code == 200:
        profile = response.json()
        full_name = profile.get('athlete', {}).get('name', 'Athlete without name')
        first_name = full_name.split()[0] if full_name else 'Athlete'
        return first_name
    logging.error(f"Error fetching athlete profile: {response.status_code}")
    return "Athlete"

athlete_name = get_athlete_name(athlete_id, username, api_key)
print(f"Athlete First Name: {athlete_name}")
logging.info(f"Using athlete first name: {athlete_name} for further processing.")

def get_wellness_data(athlete_id, username, api_key, oldest_date, newest_date):
    url_wellness = f"https://intervals.icu/api/v1/athlete/{athlete_id}/wellness"
    response = call_with_retries(requests.get, url_wellness, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        data = response.json()
        df = pd.DataFrame(data)
        df['date'] = pd.to_datetime(df['id'], errors='coerce')
        mask = (df['date'] >= pd.to_datetime(oldest_date)) & (df['date'] <= pd.to_datetime(newest_date))
        logging.info(f"Fetched wellness data for athlete {athlete_id}")
        return df.loc[mask, ['id', 'ctlLoad', 'atlLoad']].copy()
    logging.error(f"Error fetching wellness data: {response.status_code}")
    return pd.DataFrame(columns=['id', 'ctlLoad', 'atlLoad'])

def calculate_weekly_loads_vectorized(wellness_df):
    wellness_df['date'] = pd.to_datetime(wellness_df['id'], errors='coerce')
    wellness_df['year_week'] = wellness_df['date'].dt.strftime("%Y-%U")
    weekly = wellness_df.groupby('year_week')[['ctlLoad', 'atlLoad']].sum()
    logging.info("Calculated weekly loads from wellness data")
    return weekly

def get_existing_note_events(athlete_id, username, api_key, oldest_date, newest_date, prefix):
    url_get = f"{url_base}/eventsjson"
    params = {"oldest": oldest_date, "newest": newest_date, "category": "NOTE"}
    response = call_with_retries(requests.get, url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        events = response.json()
        logging.info(f"Fetched existing NOTE events for athlete {athlete_id}")
        return {ev['name']: ev for ev in events if ev['name'].startswith(prefix)}
    logging.error(f"Failed to fetch existing NOTE events: {response.status_code}")
    return {}

def delete_note_event(event_id, athlete_id, username, api_key):
    url_del = f"{url_base}/events/{event_id}"
    response_del = call_with_retries(requests.delete, url_del, headers=API_headers, auth=HTTPBasicAuth(username, api_key))
    if response_del.status_code == 200:
        logging.info(f"Deleted NOTE event ID={event_id}")
    else:
        logging.error(f"Error deleting NOTE event ID={event_id}: {response_del.status_code}")
    time.sleep(parse_delay)

def create_note_event(start_date, description, color, athlete_id, username, api_key, current_week):
    end_date = start_date
    note_ATP_name = f"{note_name_prefix_ATP} for week {current_week}"
    post_data = {
        "category": "NOTE",
        "start_date_local": start_date,
        "end_date_local": end_date,
        "name": note_ATP_name,
        "description": description,
        "not_on_fitness_chart": "true",
        "show_as_note": "false",
        "show_on_ctl_line": "false",
        "athlete_cannot_edit": "false",
        "color": color,
        "for_week": "true"
    }
    url_post = f"{url_base}/events"
    response_post = call_with_retries(requests.post, url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
    if response_post.status_code == 200:
        logging.info(f"Created NOTE event: {note_ATP_name}")
    else:
        logging.error(f"Error creating NOTE event: {note_ATP_name}, code={response_post.status_code}")
    time.sleep(parse_delay)

def update_note_event(event_id, start_date, description, color, athlete_id, username, api_key, current_week):
    end_date = start_date
    note_ATP_name = f"{note_name_prefix_ATP} for week {current_week}"
    put_data = {
        "category": "NOTE",
        "start_date_local": start_date,
        "end_date_local": end_date,
        "name": note_ATP_name,
        "description": description,
        "not_on_fitness_chart": "true",
        "show_as_note": "false",
        "show_on_ctl_line": "false",
        "athlete_cannot_edit": "false",
        "color": color,
        "for_week": "true"
    }
    url_put = f"{url_base}/events/{event_id}"
    response_put = call_with_retries(requests.put, url_put, headers=API_headers, json=put_data, auth=HTTPBasicAuth(username, api_key))
    if response_put.status_code == 200:
        logging.info(f"Updated NOTE event: {note_ATP_name}")
    else:
        logging.error(f"Error updating NOTE event: {note_ATP_name}, code={response_put.status_code}")
    time.sleep(parse_delay)

def get_first_a_event(df, note_event_date):
    note_date = datetime.strptime(note_event_date, "%Y-%m-%dT00:00:00")
    filtered = df[(pd.to_datetime(df['start_date_local']) > note_date) & (df['cat'].astype(str).str.upper() == 'A') & (df['race'].astype(str).str.strip() != '')]
    return filtered['race'].iloc[0].strip() if not filtered.empty else None

def format_focus_items_notes(focus_items_notes):
    if len(focus_items_notes) > 1:
        return ', '.join(focus_items_notes[:-1]) + ' and ' + focus_items_notes[-1]
    return ''.join(focus_items_notes)

def populate_description(description, first_a_event):
    if not description:
        description = "Nothing to mention this week."
    if first_a_event:
        description = f"- This (part) of the plan aims for **{first_a_event}**.\n\n" + description
    description = f"Hi **{athlete_name}**, here is your weekly ATP summary:\n\n" + description
    description += note_underline_ATP
    return description

def add_period_description(row, description):
    period = row['period'] if not pd.isna(row['period']) else ""
    if period:
        period_name = handle_period_name(period)
        # Use the new 'week' column from ATP_Data to add contextual info about the week in the period.
        week_val = row.get('week', None)
        week_int = None
        try:
            if week_val is not None and not pd.isna(week_val):
                week_int = int(week_val)
                # treat 0 as "no week number"
                if week_int == 0:
                    week_int = None
        except Exception:
            week_int = None

        # Get the aimed weekly load from the 'Total_load_target' column.
        weekly_target_val = None
        if 'Total_load_target' in row and not pd.isna(row.get('Total_load_target')):
            try:
                weekly_target_val = int(round(float(row.get('Total_load_target'))))
            except Exception:
                weekly_target_val = None

        # Special meaning_core for Race and Transition periods
        if period_name == "Race":
            default_meaning_core = "a focus on the upcoming race, where we prioritise tapering, sharpening and optimal rest to peak for competition"
        elif period_name == "Transition":
            default_meaning_core = "a **more easy period**, focused on recovery and consolidating training adaptations"
        elif period_name == "Peak":
            default_meaning_core = "a **Peak period**  focused on balancing load and recovery to achieve optimal race readiness (generally 1–2 weeks before the event)."        
        else:
            default_meaning_core = None

        if week_int is not None:
            # If special period types, use their default core meaning regardless of week number
            if default_meaning_core:
                meaning_core = default_meaning_core
            else:
                if week_int == 1:
                    meaning_core = "the **start week** op de trainingperiod, where we"
                elif week_int == 2:
                    meaning_core = "the **second week** op de trainingperiod, where we"
                elif week_int == 3:
                    meaning_core = "the **third week** op de trainingperiod, where we"
                elif week_int == 4:
                    meaning_core = "we ease a bit, so we just"
                else:
                    meaning_core = f"week {week_int} of the period"

            if weekly_target_val is not None:
                meaning = f"{meaning_core} aim for a TSS of **{weekly_target_val}**"
            else:
                meaning = meaning_core

            description += f"- This is **week {week_int}** of the **{period_name}** period, which means {meaning}.\n\n"
        else:
            # No week number (or zero) — different sentence form requested
            if default_meaning_core:
                meaning_core = default_meaning_core
            else:
                meaning_core = f"the **{period_name}** period"

            if weekly_target_val is not None:
                meaning = f"{meaning_core} where we aim for a TSS of **{weekly_target_val}**"
            else:
                meaning = meaning_core

            description += f"- This is **the {period_name} period**, which means {meaning}.\n\n"

        if period == "Rest":
            description += f"**{do_at_rest}**\n\n"
    return description

def add_test_description(row, description):
    test = row['test'] if not pd.isna(row['test']) else ""
    if test:
        description += f"- Do the following test(s) this week: **{test}**.\n\n"
    return description

def add_focus_description(row, description):
    focus_columns = [
        'Weight Lifting', 'Aerobic Endurance', 'Muscular force', 'Speed Skills',
        'Muscular Endurance', 'Anaerobic Endurance', 'Sprint Power'
    ]
    additional_focus = [(col, int(row.get(col, 0))) for col in focus_columns if int(row.get(col, 0)) > 0]
    additional_focus.sort(key=lambda x: x[1])
    if additional_focus:
        formatted_focus = format_focus_items_notes([col for col, _ in additional_focus])
        description += f"- Focus on **{formatted_focus}**.\n\n"
    elif description.strip():
        description += "- You don't have to focus on specific workouts this week.\n\n"
    return description

def add_race_focus_description(row, description):
    race_cat = str(row.get('cat', '')).upper()
    race_name = row.get('race', '').strip()
    if race_cat == 'A' and race_name:
        description += f"- **{race_name}** is your main-goal! This is your **{race_cat}-event**, so primarily focus on this race.\n\n"
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
            description += f"- Upcoming race: **{next_race_name}** (a **{next_race_cat}**-event) next week on {next_race_day} {next_race_dayofmonth} {next_race_month}.\n\n "
        if weeks_to_go > 1:
            description += f"- Upcoming race: **{next_race_name}** (a **{next_race_cat}**-event) within **{weeks_to_go}** weeks on {next_race_day} {next_race_dayofmonth} {next_race_month}.\n\n "
    return description

def handle_period_name(period):
    period = period.strip()
    if period == "Trans":
        return "Transition"
    elif period == "Prep":
        return "Preparation"
    elif period and not period[-1].isdigit():
        return period.strip()
    return period

def main():
    oldest_date, newest_date = read_ATP_period(ATP_file_path, sheet_name=ATP_sheet_Conditions)
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], errors='coerce')
    df = df.dropna(subset=['start_date_local'])
    oldest = pd.to_datetime(oldest_date)
    newest = pd.to_datetime(newest_date)
    df = df[(df['start_date_local'] >= oldest) & (df['start_date_local'] <= newest)]
    df['year_week'] = df['start_date_local'].apply(lambda x: f"{x.isocalendar()[0]}-{x.isocalendar()[1]}")

    logging.info("Starting ATP NOTE event sync process.")

    # Batch fetch existing NOTE events
    existing_notes = get_existing_note_events(
        athlete_id, username, api_key,
        oldest_date,
        newest_date,
        note_name_prefix_ATP
    )

    # Batch fetch wellness data and calculate weekly loads
    wellness_df = get_wellness_data(
        athlete_id, username, api_key,
        oldest_date, newest_date
    )
    weekly_loads = calculate_weekly_loads_vectorized(wellness_df)

    # Main: create or update NOTE events per week
    for index, row in df.iterrows():
        start_date = row['start_date_local'].strftime("%Y-%m-%dT00:00:00")
        week = row['start_date_local'].isocalendar()[1]
        year = row['start_date_local'].isocalendar()[0]
        note_name = f"{note_name_prefix_ATP} for week {week}"

        first_a_event = get_first_a_event(df, start_date)
        description = ""
        description = add_period_description(row, description)
        description = add_test_description(row, description)
        description = add_focus_description(row, description)
        race_focus_description = add_race_focus_description(row, description)
        if race_focus_description == description:
            description = add_next_race_description(index, df, week, description)
        else:
            description = race_focus_description

        desc_full = populate_description(description, first_a_event)

        existing_note = existing_notes.get(note_name)

        # Always create or update notes, even if nothing to mention
        if existing_note:
            # Only update if content is different
            if existing_note['description'] != desc_full:
                logging.info(f"Updating NOTE event for week {week}")
                update_note_event(existing_note['id'], start_date, desc_full, note_color_ATP, athlete_id, username, api_key, week)
            else:
                logging.info(f"No NOTE update needed for week {week}")
        else:
            logging.info(f"Creating new NOTE event for week {week}")
            create_note_event(start_date, desc_full, note_color_ATP, athlete_id, username, api_key, week)
        time.sleep(parse_delay)

if __name__ == "__main__":
    main()
