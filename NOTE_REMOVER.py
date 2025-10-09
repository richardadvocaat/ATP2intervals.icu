import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta
import xlwings as xw
import time
import random
from functools import wraps
import os
import argparse
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Configuration ---
athlete_TLA = "RAA"  # Three letter Acronym of athlete.
ATP_year = "2025"    # Year of the ATP
parse_delay = .00
coach_name = "CozyCoach"
rip_word = "Weekly"


ATP_sheet_name = "ATP_Data"
ATP_sheet_Conditions = "ATP_Conditions"
ATP_file_path = rf"C:\TEMP\{athlete_TLA}\ATP2intervals_{athlete_TLA}_{ATP_year}.xlsm"
ATP_loadcheck_sheet_name = "WTL"  # "Weekly Type Loads"
ATP_loadcheck_compare_sheet_name = "WLC"  # "Weekly Load Compare"
ATP_loadcheck_file_path = ATP_file_path   # Now writing directly to the macro file!

compliance_treshold = 0.3
note_underline_ATP = f"\n---\n *made with the {os.path.basename(__file__)} script / From coach {coach_name}*"
note_underline_FEEDBACK = f"\n---\n *made with the {os.path.basename(__file__)} script / From coach {coach_name}*"
note_underline_PERIOD = f"\n---\n *made with the {os.path.basename(__file__)} script / From coach {coach_name}*"
note_name_prefix_ATP = "Weekly training and focus summary of your ATP"
note_name_PERIOD = 'Period:'
note_name_template_FEEDBACK = "Weekly feedback about your trainingload in week {last_week}"

change_whole_range = True  # Control whether to change the whole range or only upcoming targets

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")
unit_preference = user_data.get('DISTANCE_SYSTEM', "metric")
note_color_ATP = user_data.get('NOTE_ATP_COLOR', "red")
note_color_FEEDBACK = user_data.get('NOTE_FEEDBACK_COLOR', "blue")
do_at_rest = user_data.get('Do_At_Rest', "Do nothing!")

url_base = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"{url_base}/profile"
url_activities = f"{url_base}/activities"
API_headers = {"Content-Type": "application/json"}
def delete_note_events(year, rip_word, verbose=False):
    """Deletes NOTE events containing `rip_word` for the specified year."""
    start_date = datetime(year, 1, 1).strftime("%Y-%m-%dT00:00:00")
    end_date = datetime(year, 12, 31).strftime("%Y-%m-%dT23:59:59")
    url_get = f"{url_base}/events.json"
    params = {"oldest": start_date, "newest": end_date, "category": "NOTE"}
    headers = API_headers

    try:
        resp = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(username, api_key))
        resp.raise_for_status()
        events = resp.json()
        if verbose:
            logging.info(f"Fetched {len(events)} NOTE events for {year}")
        deleted = 0
        for event in events:
            if rip_word.lower() in event['name'].lower():
                event_id = event['id']
                url_del = f"{url_base}/events/{event_id}"
                del_resp = requests.delete(url_del, headers=headers, auth=HTTPBasicAuth(username, api_key))
                if del_resp.ok:
                    deleted += 1
                    logging.info(f"Deleted event ID={event_id} - Name: {event['name']}")
                else:
                    logging.error(f"Failed to delete event ID={event_id} - Status: {del_resp.status_code}")
        print(f"Deleted {deleted} events containing '{rip_word}' for year {year}.")
    except Exception as e:
        logging.error(f"Error processing notes: {e}")
        print("Failed to process/delete events.")

def main():
    parser = argparse.ArgumentParser(description="Delete NOTES containing a specific word for a given year.")
    parser.add_argument("--year", type=int, help="Year to check (e.g., 2026). If not provided, prompts interactively.")
    parser.add_argument("--rip_word", type=str, help="Word to match in NOTES to delete. If not provided, prompts interactively.")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging.")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO if args.verbose else logging.WARNING, format='%(asctime)s - %(levelname)s - %(message)s')

    # Prompt interactively if not supplied
    year = args.year if args.year else int(input("Year to check for NOTE events to delete? "))
    rip_word = args.rip_word if args.rip_word else input("Word to search for in NOTE events to delete (rip_word)? ")

    delete_note_events(year, rip_word, args.verbose)

if __name__ == "__main__":
    main()
