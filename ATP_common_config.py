import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta
import xlwings as xw

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Configuration ---
athlete_TLA = "TLA"  # Three letter Acronym of athlete.
ATP_year = "YYYY"    # Year of the ATP
parse_delay = .00
coach_name = "CozyCoach"

ATP_sheet_name = "ATP_Data"
ATP_sheet_Conditions = "ATP_Conditions"
ATP_file_path = rf"C:\TEMP\{athlete_TLA}\ATP2intervals_{athlete_TLA}_{ATP_year}.xlsm"
ATP_loadcheck_sheet_name = "WTL"  # "Weekly Type Loads"
ATP_loadcheck_compare_sheet_name = "WLC"  # "Weekly Load Compare"
ATP_loadcheck_file_path = ATP_file_path   # Now writing directly to the macro file!

compliance_treshold = 0.3

note_underline_ATP = f"\n---\n *made with the 2_ATP_NOTES.py script / From coach {coach_name}*"
note_underline_FEEDBACK = f"\n---\n *made with the 5_ATP_WEEKLY_LOAD_FEEDBACK_NOTES.py / From coach {coach_name}*"
note_underline_PERIOD = f"\n---\n *made with the 3_ATP_PERIOD_NOTE.py script / From coach {coach_name}*"
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
