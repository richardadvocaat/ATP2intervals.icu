import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

Athlete_TLA = "RAA" #Three letter Acronym of athlete.
ATP_sheet_name = "ATP_Data"
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\Intervals_API_Tools_Office365_v1.6_ATP2intervals_{Athlete_TLA}.xlsm"

parse_delay = .01
note_ATP_name = "Weekly training and focus summary of your ATP"

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")
unit_preference = user_data.get('DISTANCE_SYSTEM', "metric")
note_ATP_color = user_data.get('NOTE_ATP_COLOR', "red")
note_FEEDBACK_color = user_data.get('NOTE_FEEDBACK_COLOR', "blue")
do_at_rest = user_data.get('Do_At_Rest', "Do nothing!")

url_base = "https://intervals.icu/api/v1/athlete/{athlete_id}"
url_profile = f"https://intervals.icu/api/v1/athlete/{athlete_id}/profile"
url_activities = f"https://intervals.icu/api/v1/athlete/{athlete_id}/activities"
API_headers = {"Content-Type": "application/json"}

def get_last_day_of_week(date):
    return date + timedelta(days=(6 - date.weekday()))

def get_period_end_date(df, start_index):
    period = df.at[start_index, 'period']
    for i in range(start_index, len(df)):
        if df.at[i, 'period'] != period:
            return get_last_day_of_week(df.at[i-1, 'start_date_local'])
    return get_last_day_of_week(df.at[len(df)-1, 'start_date_local'])

def get_note_color(period):
    color_mapping = {
        "Base": "yellow",
        "Peak": "orange",
        "Race": "red",
        "Trans": "green",
        "Prep": "blue",
        "Recovery": "purple",
        "Rest": "cyan"
    }
    return color_mapping.get(period, "black")  # Default to black if period not found

def create_note_event(start_date, end_date, description, period, athlete_id, username, api_key):
    url_base = f"https://intervals.icu/api/v1/athlete/{athlete_id}"
    url_post = f"{url_base}/events"
    API_headers = {"Content-Type": "application/json"}
    
    period_full = period
    if period == "Trans":
        period_full = "Transition"
    elif period == "Prep":
        period_full = "Preparation"
    
    color = get_note_color(period)
    
    post_data = {
        "category": "NOTE",
        "start_date_local": start_date.strftime("%Y-%m-%dT00:00:00"),
        "end_date_local": end_date.strftime("%Y-%m-%dT00:00:00"),
        "name": f"Training Period: {period_full}",
        "description": description,
        "color": color
    }
    
    response_post = requests.post(url_post, headers=API_headers, json=post_data, auth=HTTPBasicAuth(username, api_key))
    if response_post.status_code == 200:
        logging.info(f"New event created from {start_date} to {end_date}!")
    else:
        logging.error(f"Error creating event: {response_post.status_code}")

def main():
    
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df['start_date_local'] = pd.to_datetime(df['start_date_local'], format='%d-%b')
    df.fillna('', inplace=True)
    
    for i in range(len(df)):
        start_date = df.at[i, 'start_date_local']
        period = df.at[i, 'period']
        
        if i == 0 or df.at[i-1, 'period'] != period:
            end_date = get_period_end_date(df, i)
            description = f"You are in the **{period}-period** (Which goes from {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')})"
            create_note_event(start_date, end_date, description, period, athlete_id, username, api_key)

if __name__ == "__main__":
    main()
