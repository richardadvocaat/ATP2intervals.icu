import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta
from io import StringIO
import xlwings as xw  # Added xlwings import

def format_activity_name(activity):
    return ''.join(word.capitalize() for word in activity.split('_'))

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

Athlete_TLA = "RAA"
ATP_year = "2026"
ATP_sheet_name = "ATP_Data"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\ATP2intervals_{Athlete_TLA}_{ATP_year}.xlsm"
RACE_file_path = rf"C:\TEMP\{Athlete_TLA}\{Athlete_TLA}_race_events_{ATP_year}.xlsx"

parse_delay = .01

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")

url_base = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
url_activities = f"https://intervals.icu/api/v1/athlete/{athlete_id}/activities"
API_headers = {"Content-Type": "application/json"}

oldest_date = f"{ATP_year}-01-01T00:00:00"
newest_date = f"{ATP_year}-12-31T23:59:59"

def get_race_events(athlete_id, username, api_key, race_categories, oldest_date, newest_date):
    url_get = f"https://intervals.icu/api/v1/athlete/{athlete_id}/eventsjson"
    events_list = []
    for category in race_categories:
        params = {
            "oldest": oldest_date,
            "newest": newest_date,
            "category": category
        }
        logging.info(f"Requesting {url_get} with params: {params}")
        response = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
        logging.info(f"Response Status: {response.status_code}, Body: {response.text[:500]}")
        if response.status_code == 200:
            try:
                events = response.json()  # parse as JSON
                logging.info(f"Fetched events for category {category}: {len(events)} events")
                events_list.append((category, events))
            except Exception as e:
                logging.error(f"Error parsing JSON for category {category}: {e}")
        else:
            logging.error(f"Error fetching events for category {category}: {response.status_code}. {response.text}")
    return events_list

def structurize_events_json(events_json):
    if not events_json:
        return pd.DataFrame()
    df = pd.DataFrame(events_json)
    # Select and rename required columns
    cols_map = {
        "type": "racetype",
        "category": "racecategory",
        "end_date_local": "date",
        "name": "racename"
    }
    df = df[list(cols_map.keys())].rename(columns=cols_map)
    # Convert date to DD-MM-YYYY format
    df["date"] = pd.to_datetime(df["date"]).dt.strftime("%d-%m-%Y")
    # Reorder columns as requested
    df = df[["date", "racename", "racetype", "racecategory"]]
    return df

def save_events_to_excel(race_events, output_file):
    app = xw.App(visible=False)
    try:
        wb = xw.Book()
        for category, events in race_events:
            try:
                df = structurize_events_json(events)
                if not df.empty:
                    if category in [s.name for s in wb.sheets]:
                        wb.sheets[category].delete()
                    sht = wb.sheets.add(category)
                    sht.range("A1").value = df.columns.tolist()
                    sht.range("A2").value = df.values.tolist()
                else:
                    logging.warning(f"No events to save for category {category}")
            except Exception as e:
                logging.error(f"Error processing events for category {category}: {e}")
        if "Sheet1" in [s.name for s in wb.sheets] and len(wb.sheets) > len(race_events):
            wb.sheets["Sheet1"].delete()
        wb.save(output_file)
        logging.info(f"Race events have been saved to {output_file}")
    finally:
        wb.close()
        app.quit()

# --------------- ADDED WORKFLOW: CHECK RACE CALENDAR ---------------

def read_race_calendar_table(file_path, sheet_name="Race_Calendar", start_cell="B2"):
    app = xw.App(visible=False)
    try:
        wb = xw.Book(ATP_file_path)
        sht = wb.sheets[sheet_name]
        headers = sht.range(start_cell).expand('right').value
        data = sht.range(f"{start_cell}:{start_cell}").offset(1, 0).expand('table').value
        df = pd.DataFrame(data, columns=headers)
        return df, wb, sht
    finally:
        app.quit()

def get_intervals_races(race_events):
    all_events = []
    for cat, events in race_events:
        for event in events:
            # Use category from Intervals.icu event if present, else fall back to request category
            event_cat = event.get("category", cat)
            event_name = event.get("name")
            event_date_local = event.get("end_date_local")
            try:
                # Try extracting date as date object if possible
                event_date = pd.to_datetime(event_date_local).date() if event_date_local else None
            except Exception:
                event_date = None
            all_events.append({
                "racename": event_name,
                "date": event_date,
                "racecategory": event_cat
            })
    return pd.DataFrame(all_events)

def check_in_intervals(race_calendar_df, intervals_df):
    result = race_calendar_df.copy()
    result["in_intervals?"] = "no"
    result["cat_correct?"] = "no"
    for idx, row in result.iterrows():
        race_name = str(row["race_name"]).strip()
        race_date_key = row.get("race_date_key", None)
        # Try parsing date from race_date_key, fallback to race_date
        try:
            if race_date_key:
                race_date = pd.to_datetime(race_date_key).date()
            elif "race_date" in row and row["race_date"]:
                race_date = pd.to_datetime(row["race_date"]).date()
            else:
                race_date = None
        except Exception:
            race_date = None
        race_cat = str(row.get("race_cat", "")).strip()
        # Match on race name and date
        match = intervals_df[
            (intervals_df["racename"].fillna("").str.strip() == race_name) &
            (intervals_df["date"] == race_date)
        ]
        if not match.empty:
            result.at[idx, "in_intervals?"] = "yes"
            # Check category correctness
            if race_cat in match["racecategory"].astype(str).str.strip().values:
                result.at[idx, "cat_correct?"] = "yes"
    return result

def write_table_to_excel(result_df, wb, sht, start_cell="B2"):
    # Overwrite original table with new columns
    sht.range(start_cell).value = result_df.columns.tolist()
    sht.range(start_cell).offset(1,0).value = result_df.values.tolist()
    wb.save()

# --------------- MAIN ---------------

def main():
    race_categories = ["RACE_A", "RACE_B", "RACE_C"]
    race_events = get_race_events(athlete_id, username, api_key, race_categories, oldest_date, newest_date)
    if race_events:
        output_file = RACE_file_path
        save_events_to_excel(race_events, output_file)
        print(f"Race events have been saved to {output_file}")
        
        # --- ADDED: Race_Calendar sheet checking ---
        try:
            race_calendar_df, wb, sht = read_race_calendar_table(ATP_file_path, "Race_Calendar", "B2")
        except Exception as e:
            print(f"Could not read Race_Calendar sheet: {e}")
            return
        intervals_df = get_intervals_races(race_events)
        result_df = check_in_intervals(race_calendar_df, intervals_df)
        write_table_to_excel(result_df, wb, sht, "B2")
        print("Race_Calendar sheet updated with Intervals.icu match info.")
    else:
        print("No RACE events found.")

if __name__ == "__main__":
    main()
