# ATP2intervals.icu

To Do: Store parameters for the coach in a separate document where also the coached athlete names are stored. (or get this list from intervals.icu.)
The idea is that it must be an easy job to populate intervals.icu with the coach's made ATP.

## What's in it
**The first Python script ([1_ATP_populate.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/1_ATP_populate.py))** automates the process of sending an annual training plan (ATP) with weekly (TSS, Time and/or distance)-loads for multiple sports. It reads data from the ATP Excel file made by [Gerald](https://forum.intervals.icu/u/gerald/summary) These files can be found on [Google Drive](https://drive.google.com/drive/folders/1WhIOf2XkGiZBEN_7tX2PSShmF-QXBnBF).
Find more information about ATP on the [intervals.icu forum](https://forum.intervals.icu/t/apps-in-excel-a-guide-to-getting-started/20844).

**The second script ([2_ATP_NOTES.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/2_ATP_NOTES.py))** makes notes with multiple comments for your athlete to intervals.icu. It reads data like focus on "Most Important Workouts", tests and coming events from the ATP Excel file made by [Gerald](https://forum.intervals.icu/u/gerald/summary) 

**The third script ([3_ATP_LOAD_feedback.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/3_ATP_LOAD_feedback.py))** checks if the athlete was compliant to the plan based on the total load of the prior week compared with the load of the ATL. It gives some basic feedback based on the difference in load.

**The fourth script ([4_LOAD_CHECK.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/LOAD_CHECK/4_LOAD_CHECK.py))** compares the planned target loads in intervals.icu with the planned loads from the workouts. This give you as a trainer/coach an easy way to check if the WORKOUTS in intervals.icu are in line with the TARGET_LOADS from the ATP. The script makes a sheet called  ATP_LOAD.xlsx and puts in in C:\TEMP. The main sheet looks for the data in this sheet. (direct writing the data in the main sheet corrupts it...needs no be fixed.)

**The fifth script ([5_ATP_PERIOD_NOTE.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/5_ATP_PERIOD_NOTE.py))** makes a note along the whole period (build, transition, race etc.). So it is visible along the the period and also on the fittness chart.

## Features

- **Read Excel File**: Reads an Excel file containing weekly load, time, and distance targets for different activities (Bike, Run, Swim).
- **Create or Update Events**: Creates new events or updates existing ones in intervals.icu based on load targets.
- **Delete Events**: Deletes corresponding events in intervals.icu if the load target is zero.
- **Efficient Data Handling**: Retrieves all relevant events in a single API call to improve performance and reduce server load.
- **Unit Conversion**: Allows users to specify their unit preference (metric or imperial) for Bike and Run distances, while Swim distances remain in meters.
- **Test Column**: Adds comments for tests specified in the 'test' column.
- **Focus Columns**: Adds focus based on specified focus columns like Aerobic Endurance, Muscular Force, etc.
- **Race Categories**: Adds focus for A, B, and C category races with custom descriptions.
- **Fetch Athlete Name**: Fetches and uses the athlete's first name from intervals.icu for personalized messages.
- **Event Management**: Functions to create, update, or delete target and note events.
- **Description Handling**: Functions to handle different parts of the description, such as period, test, focus, race focus, and next race descriptions.
- **Separated Functions**:
  - **Load Targets**: Uploads load targets and adds notes based on the data from the ATP.
  - **Load vs Target Load**: Handles notes related to the loads vs target load in the week before.
- **Distance Conversion**: Converts distances for different units (metric, imperial).
- **Weekly Wellness Data**: Retrieves and calculates weekly wellness data (ctlLoad and atlLoad).
- **Training Load Feedback**: Provides feedback on the athlete's training load for the previous week.
- **Event Prefix Deletion**: Deletes events based on a prefix.
- **Excel Export**: Exports data to Excel with planned loads and target loads in separate sheets, and sets column widths based on the header/name of the column.

## Usage

1. **Setup**: Ensure you have the required Python libraries installed (`logging`, `os`, `pandas`, `requests`).
2. **Configuration**: Update the user variables at the top of the script with your Excel file path, sheet name.
3. **Excel Sheets**: Place the `Intervals_API_Tools_Office365_v1._ATP2intervals_TLA.xlsm` in `C:\TEMP\TLA`. (TLA = e.g. RAA, so rename the sheet to Intervals_API_Tools_Office365_v1._ATP2intervals_RAA.xlsm  and put it in 'C:\TEMP\RAA' )
4. **User_Data**: In the tab User_Data you can put the athlete ID, API key, preferred unitsystem (metric or imperial) and some basic preferences like the color of the note.
5. **Run**: Execute the scripts in the right order to sync your training plan with intervals.icu.


