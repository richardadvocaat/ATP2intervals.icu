# ATP2intervals.icu

ISSUE: DATE of UPCOMING RACE DOES NOT MATCH IN THE DECRIPTION
To Do: make path to the ATP file dependend of the athlete's name. Store parameters for the trainer in a separate document where also the athletes are stored.(or get this list from intervals.
at the end it must be fair easy job to populate intervals.icu with the coach's made ATP quite easily.

The first Python script ([ATP2Intervals.py_ATP_populate.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/richardadvocaat-patch-1-LOAD-FEADBACK/ATP2Intervals.py_ATP_populate.py)) automates the process of sending an annual training plan (ATP) with weekly (TSS, Time and/or distance)-loads for multiple sports. And it makes notes with multiple comments for your athlete to intervals.icu. It reads data from the ATP Excel file made by [Gerald](https://forum.intervals.icu/u/gerald/summary) These files can be found on [Google Drive](https://drive.google.com/drive/folders/1WhIOf2XkGiZBEN_7tX2PSShmF-QXBnBF).
Find more information about ATP on the [intervals.icu forum](https://forum.intervals.icu/t/apps-in-excel-a-guide-to-getting-started/20844).

The second script ([ATP2Intervals.py_LOAD_FEEDBACK.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/richardadvocaat-patch-1-LOAD-FEADBACK/ATP2Intervals.py_LOAD_FEEDBACK.py) checks if the athlete was compliant to the plan based on the total load of the prior week compared with the load of the ATL. It gives some basic feedback.


## Features

- **Read Excel File**: Reads an Excel file containing weekly load, time, and distance targets for different activities (Bike, Run, Swim).
- **Create or Update Events**: Creates new events or updates existing ones in intervals.icu based on load targets.
- **Delete Events**: Deletes corresponding events in intervals.icu if the load target is zero.
- **Efficient Data Handling**: Retrieves all relevant events in a single API call to improve performance and reduce server load.
- **Unit Conversion**: Allows users to specify their unit preference (metric or imperial) for Bike and Run distances, while Swim distances remain in meters.
- **Custom Descriptions**: Adds custom descriptions based on the 'period' and 'focus' columns from the Excel file.
  - If 'period' is "Rest", the description will include something like "Stay in bed or on the beach!". (but it is a variable)
  - If 'focus' has a value, the description will include "Focus this week on {focus}".
- **Test Column**: Adds comments for tests specified in the 'test' column.
- **Focus Columns**: Adds focus based on specified focus columns like Aerobic Endurance, Muscular Force, etc.
- **Race Categories**: Adds focus for A, B, and C category races with custom descriptions.
- **Fetch Athlete Name**: Fetches and uses the athlete's first name from intervals.icu for personalized messages.
- **Event Management**: Functions to create, update, or delete target and note events.
- **Description Handling**: Functions to handle different parts of the description, such as period, test, focus, race focus, and next race descriptions.
- **Separated Functions**:
  - **Load Targets**: Uploads load targets and adds notes based on the data from the ATP.
  - **Load vs Target Load**: Handles notes related to the loads vs target load in the week before.

## Usage

1. **Setup**: Ensure you have the required Python libraries installed (`logging`, `os`, `pandas`, `requests`).
2. **Configuration**: Update the user variables at the top of the script with your Excel file path, sheet name, athlete ID, API key, and preferred unit system (metric or imperial).
3. **Excel Sheets**: Place the `Intervals_API_Tools_Office365_v1._ATP2intervals.xlsm` in `C:\TEMP`. (This is hardcoded now, but based on the name, there is a possibility to differentiate this)
4. **Run**: Execute the script to sync your training plan with intervals.icu.

