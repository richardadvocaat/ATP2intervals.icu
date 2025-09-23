import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime
import xlwings as xw

def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    user_data = df.set_index('Key').to_dict()['Value']
    return user_data

Athlete_TLA = "TLA" #Three letter Acronym of athlete.
ATP_year = "YYYY" #The year of the plan.
ATP_file_path = rf"C:\TEMP\{Athlete_TLA}\ATP2intervals_{Athlete_TLA}_{ATP_year}.xlsm"

ATP_sheet_name = "ATP_Data"
#ATP_loadcheck_sheet_name = "Weekly Type Loads"
ATP_loadcheck_sheet_name = "WTL"
#ATP_loadcheck_compare_sheet_name = "Weekly Load Compare"
ATP_loadcheck_compare_sheet_name = "WLC"

ATP_loadcheck_file_path = ATP_file_path   # Now writing directly to the macro file!

user_data = read_user_data(ATP_file_path)
api_key = user_data.get('API_KEY', "yourapikey")
username = user_data.get('USERNAME', "API_KEY")
athlete_id = user_data.get('ATHLETE_ID', "athleteid")

url_base = "https://intervals.icu/api/v1/athlete/{athlete_id}"
API_headers = {"Content-Type": "application/json"}

def get_events(athlete_id, username, api_key, oldest_date, newest_date, category):
    url_get = f"{url_base}/eventsjson".format(athlete_id=athlete_id)
    params = {"oldest": oldest_date.strftime("%Y-%m-%dT00:00:00"), "newest": newest_date.strftime("%Y-%m-%dT00:00:00"), "category": category}
    response = requests.get(url_get, headers=API_headers, params=params, auth=HTTPBasicAuth(username, api_key))
    if response.status_code == 200:
        return response.json()
    else:
        logging.error(f"Error fetching events for category {category}: {response.status_code}")
        return []

def calculate_weekly_type_loads(workouts, race_b_events, race_c_events):
    weekly_type_loads = {}
    for workout in workouts + race_b_events + race_c_events:
        if 'id' not in workout:
            continue
        date = datetime.strptime(workout['start_date_local'], "%Y-%m-%dT%H:%M:%S")
        week = f"{date.isocalendar()[1]:02d}"  # Format week to two digits
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
        week = f"{date.isocalendar()[1]:02d}"  # Format week to two digits
        year = date.isocalendar()[0]
        year_week = f"{year}-{week}"
        load_target = target.get('load_target') or 0
        target_type = target.get('type', 'Unknown')
        if year_week not in weekly_target_loads:
            weekly_target_loads[year_week] = {}
        if target_type not in weekly_target_loads[year_week]:
            weekly_target_loads[year_week][target_type] = 0
        weekly_target_loads[year_week][target_type] += load_target
    return weekly_target_loads

def set_column_widths(sheet, df, start_col=1):
    for i, col in enumerate(df.columns, start=start_col):
        maxlen = max(df[col].astype(str).map(len).max(), len(col))
        sheet.range((1, i)).column_width = maxlen + 2

def export_to_excel(weekly_type_loads, weekly_target_loads, file_path):
    # Create DataFrames as before (pandas is the best tool for this!)
    rows = []
    all_types = set()
    for year_week in set(weekly_type_loads.keys()).union(weekly_target_loads.keys()):
        row = {"Week": year_week}
        for workout_type in weekly_type_loads.get(year_week, {}):
            row[f"Actual {workout_type}"] = weekly_type_loads[year_week][workout_type]
            all_types.add(workout_type)
        rows.append(row)
    planned_df = pd.DataFrame(rows).fillna(0)
    actual_columns = sorted([f"Actual {t}" for t in all_types])
    for col in actual_columns:
        if col not in planned_df.columns:
            planned_df[col] = 0
    planned_df = planned_df[["Week"] + actual_columns].sort_values(by="Week")

    rows = []
    for year_week in set(weekly_type_loads.keys()).union(weekly_target_loads.keys()):
        row = {"Week": year_week}
        for workout_type in weekly_type_loads.get(year_week, {}):
            row[f"Actual {workout_type}"] = weekly_type_loads[year_week][workout_type]
        for target_type in weekly_target_loads.get(year_week, {}):
            row[f"Target {target_type}"] = weekly_target_loads[year_week][target_type]
        row["Total Actual_Load"] = sum(row.get(f"Actual {t}", 0) for t in all_types)
        row["Total Target_Load"] = sum(row.get(f"Target {t}", 0) for t in all_types)
        row["Load Difference"] = row["Total Actual_Load"] - row["Total Target_Load"]
        rows.append(row)
    compare_df = pd.DataFrame(rows).fillna(0)
    target_columns = sorted([f"Target {t}" for t in all_types])
    for col in actual_columns + target_columns:
        if col not in compare_df.columns:
            compare_df[col] = 0
    compare_df = compare_df[["Week"] + actual_columns + target_columns + ["Total Actual_Load", "Total Target_Load", "Load Difference"]].sort_values(by="Week")

    # Use xlwings to write to the .xlsm file
    app = xw.App(visible=False)
    try:
        wb = xw.Book(file_path)
        # Write planned_df to ATP_loadcheck_sheet_name (starting at row 4, col 1)
        if ATP_loadcheck_sheet_name not in [s.name for s in wb.sheets]:
            wb.sheets.add(ATP_loadcheck_sheet_name)
        ws_planned = wb.sheets[ATP_loadcheck_sheet_name]
        ws_planned.range("A4").options(index=False, header=True).value = planned_df
        set_column_widths(ws_planned, planned_df, start_col=1)

        # Write compare_df to ATP_loadcheck_compare_sheet_name (starting at row 4, col 1)
        if ATP_loadcheck_compare_sheet_name not in [s.name for s in wb.sheets]:
            wb.sheets.add(ATP_loadcheck_compare_sheet_name)
        ws_compare = wb.sheets[ATP_loadcheck_compare_sheet_name]
        ws_compare.range("A4").options(index=False, header=True).value = compare_df
        set_column_widths(ws_compare, compare_df, start_col=1)

        # Call VBA macro to create tables from the written ranges
        # Assumes your workbook has a VBA macro named CreateTableOnSheet as described before
        # Macro signature: Sub CreateTableOnSheet(sheetName As String, startCell As String)
        wb.macro('CreateTableOnSheet')(ATP_loadcheck_sheet_name, 'A4')
        wb.macro('CreateTableOnSheet')(ATP_loadcheck_compare_sheet_name, 'A4')

        wb.save()
        wb.close()
    finally:
        app.quit()

def main():
    df = pd.read_excel(ATP_file_path, sheet_name=ATP_sheet_name)
    df.fillna(0, inplace=True)
    oldest_date = df['start_date_local'].min()
    newest_date = df['start_date_local'].max()
    workouts = get_events(athlete_id, username, api_key, oldest_date, newest_date, "WORKOUT")
    race_b_events = get_events(athlete_id, username, api_key, oldest_date, newest_date, "RACE_B")
    race_c_events = get_events(athlete_id, username, api_key, oldest_date, newest_date, "RACE_C")
    target_loads = get_events(athlete_id, username, api_key, oldest_date, newest_date, "TARGET")
    weekly_type_loads = calculate_weekly_type_loads(workouts, race_b_events, race_c_events)
    weekly_target_loads = calculate_weekly_target_loads(target_loads)
    export_to_excel(weekly_type_loads, weekly_target_loads, ATP_loadcheck_file_path)

if __name__ == "__main__":
    main()
