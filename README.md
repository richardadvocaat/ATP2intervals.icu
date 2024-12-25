# ATP2intervals.icu

This Python script automates the process of sending an annual training plan with weekly loads (for multiple sports) and a load of comments for your athlete to intervals.icu. It reads data from the ATP Excel file made by [Gerald](https://forum.intervals.icu/u/gerald/summary) based on the files to be found [here](https://drive.google.com/drive/folders/1WhIOf2XkGiZBEN_7tX2PSShmF-QXBnBF).
Find information about ATP [here](https://forum.intervals.icu/t/apps-in-excel-a-guide-to-getting-started/20844) on the intervals.icu forum).

## Features

- **Read Excel File**: Reads an Excel file containing weekly load, time, and distance targets for different activities (Bike, Run, Swim).
- **Create or Update Events**: Creates new events or updates existing ones in intervals.icu based on load targets.
- **Delete Events**: Deletes corresponding events in intervals.icu if the load target is zero.
- **Efficient Data Handling**: Retrieves all relevant events in a single API call to improve performance and reduce server load.
- **Unit Conversion**: Allows users to specify their unit preference (metric or imperial) for Bike and Run distances, while Swim distances remain in meters.
- **Custom Descriptions**: Adds custom descriptions based on the 'period' and 'focus' columns from the Excel file.
  - If 'period' is "Rest", the description will include someting like "Stay in bed or on the beach!". (but it an variable)
  - If 'focus' has a value, the description will include "Focus this week on {focus}".
- **Test Column**: Adds comments for tests specified in the 'test' column.
- **Focus Columns**: Adds focus based on specified focus columns like Aerobic Endurance, Muscular Force, etc.
- **Race Categories**: Adds focus for A, B, and C category races with custom descriptions.
- **Fetch Athlete Name**: Fetches and uses the athlete's first name from intervals.icu for personalized messages.
- **Event Management**: Functions to create, update, or delete target and note events.
- **Description Handling**: Functions to handle different parts of the description, such as period, test, focus, race focus, and next race descriptions.

## Usage

1. **Setup**: Ensure you have the required Python libraries installed (`logging`, `os`, `pandas`, `requests`).
2. **Configuration**: Update the user variables at the top of the script with your Excel file path, sheet name, athlete ID, API key, and preferred unit system (metric or imperial).
3. **Excel Sheets**: Place the `USERDATA.xlsx` in `C:\TEMP`. Configure the path to `ATP.xlsx` in `USERDATA.xlsx` (default is `C:\TEMP\ATP.xlsx`).
4. **Run**: Execute the script to sync your training plan with intervals.icu.

