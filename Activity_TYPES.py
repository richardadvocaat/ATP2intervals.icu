import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

def fetch_activity_types(athlete_id, username, api_key):
    url = f"https://intervals.icu/api/v1/athlete/{athlete_id}/sport-settings"
    response = requests.get(url, auth=HTTPBasicAuth(username, api_key))

    if response.status_code == 200:
        return response.json()
    else:
        logging.error(f"Failed to fetch sport settings: {response.status_code}")
        return []

def save_to_excel(data, file_path):
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)
    logging.info(f"Data saved to {file_path}")

def main():
    
    Athlete_TLA = "RAA"
    ATP_year = "2026"
    ATP_sheet_name = "ATP_Data"
    ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\ATP2intervals_{Athlete_TLA}_{ATP_year}.xlsm"
    user_data = read_user_data(ATP_file_path)
    api_key = user_data.get('API_KEY', "yourapikey")
    username = user_data.get('USERNAME', "API_KEY")
    athlete_id = user_data.get('ATHLETE_ID', "athleteid")

    sport_settings = fetch_activity_types(athlete_id, username, api_key)
    activity_data = []

    for sport in sport_settings:
        for activity_type in sport['types']:
            activity_data.append({
                'athlete_id': sport['athlete_id'],
                'sport_id': sport['id'],
                'activity_type': activity_type
            })

    save_to_excel(activity_data, r'C:\TEMP\ACTIVITIES.xlsx')

if __name__ == "__main__":
    main()
