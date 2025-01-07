import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

ATP_sheet_name = "ATP_Data"
ATP_file_path = r'C:\TEMP\Intervals_API_Tools_Office365_v1.6_ATP2intervals.xlsm'

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")

url_base = "https://intervals.icu/api/v1/athlete/{athlete_id}"
API_headers = {"Content-Type": "application/json"}

def get_workouts(athlete_id, username, api_key, oldest_date, newest_date):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date.strftime("%Y-%m-%dT00:00:00"), "newest": newest_date.strftime("%Y-%m-%dT00:00:00"), "category": "WORKOUT"}
    response = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        return response.json()
    else:
        logging.error(f"Error fetching workouts: {response.status_code}")
        return []

def get_target_loads(athlete_id, username, api_key, oldest_date, newest_date):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date.strftime("%Y-%m-%dT00:00:00"), "newest": newest_date.strftime("%Y-%m-%dT00:00:00"), "category": "TARGET"}
    response = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        return response.json()
    else:
        logging.error(f"Error fetching target loads: {response.status_code}")
        return []

def calculate_weekly_type_loads(workouts):
    weekly_type_loads = {}

    for workout in workouts:
        if 'id' not in workout:
            continue
        date = datetime.strptime(workout['start_date_local'], "%Y-%m-%dT%H:%M:%S")
        week = date.isocalendar()[1]
        year = date.isocalendar()[0]
        year_week = f"{year}-{week}"
        
        icu_training_load = workout.get('icu_training_load') or 0
        workout_type = workout.get('type', 'Unknown')
        
        if year_week not in weekly_type_loads:
            weekly_type_loads[year_week] = {}
        if workout_type not in weekly_type_loads[year_week]:
            weekly_type_loads[year_week][workout_type] = 0
        weekly_type_loads[year_week][workout_type] += icu_training_load

    return weekly_type_loads

def calculate_weekly_target_loads(target_loads):
    weekly_target_loads = {}

    for target in target_loads:
        if 'id' not in target:
            continue
        date = datetime.strptime(target['start_date_local'], "%Y-%m-%dT%H:%M:%S")
        week = date.isocalendar()[1]
        year = date.isocalendar()[0]
        year_week = f"{year}-{week}"
        
        target_load = target.get('target_load') or 0
        target_type = target.get('type', 'Unknown')
        
        if year_week not in weekly_target_loads:
            weekly_target_loads[year_week] = {}
        if target_type not in weekly_target_loads[year_week]:
            weekly_target_loads[year_week][target_type] = 0
        weekly_target_loads[year_week][target_type] += target_load

    return weekly_target_loads

def export_to_excel(weekly_type_loads, weekly_target_loads, file_path):
    rows = []
    all_types = set()

    for year_week in set(weekly_type_loads.keys()).union(weekly_target_loads.keys()):
        row = {"Week": year_week}
        for workout_type in weekly_type_loads.get(year_week, {}):
            row[f"Actual {workout_type}"] = weekly_type_loads[year_week][workout_type]
            all_types.add(workout_type)
        for target_type in weekly_target_loads.get(year_week, {}):
            row[f"Target {target_type}"] = weekly_target_loads[year_week][target_type]
            all_types.add(target_type)
        rows.append(row)

    df = pd.DataFrame(rows).fillna(0)
    actual_columns = sorted([f"Actual {t}" for t in all_types])
    target_columns = sorted([f"Target {t}" for t in all_types])
    # Ensure all columns are present in the DataFrame
    for col in actual_columns + target_columns:
        if col not in df.columns:
            df[col] = 0
    df = df[["Week"] + actual_columns + target_columns]

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Weekly Type Loads', index=False)

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)

    oldest_date = df['start_date_local'].min()
    newest_date = df['start_date_local'].max()

    workouts = get_workouts(athlete_id, username, api_key, oldest_date, newest_date)
    target_loads = get_target_loads(athlete_id, username, api_key, oldest_date, newest_date)
    weekly_type_loads = calculate_weekly_type_loads(workouts)
    weekly_target_loads = calculate_weekly_target_loads(target_loads)

    export_to_excel(weekly_type_loads, weekly_target_loads, r'C:\TEMP\ATP_LOAD.xlsx')

if __name__ == "__main__":
    main()
