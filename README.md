# ATP2intervals.icu

To Do: Store parameters for the coach in a separate document where also the coached athlete names are stored. (or get this list from intervals.icu.)
The idea is that it must be an easy job to populate intervals.icu with the coach's made ATP.

## What's in it
**The first Python script ([1_ATP_populate.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/1_ATP_populate.py))** automates the process of sending an annual training plan (ATP) with the coach's made ATP to intervals.icu.
Find more information about ATP on the [intervals.icu forum](https://forum.intervals.icu/t/apps-in-excel-a-guide-to-getting-started/20844).

**The second script ([2_ATP_NOTES.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/2_ATP_NOTES.py))** makes notes with multiple comments for your athlete to intervals.icu. It reads data from the ATP and sends notes to intervals.icu.

**The third script ([3_ATP_LOAD_feedback.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/3_ATP_LOAD_feedback.py))** checks if the athlete was compliant to the plan based on the reported load, time, and distance in intervals.icu.

**The fourth script ([4_LOAD_CHECK.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/LOAD_CHECK/4_LOAD_CHECK.py))** compares the planned target loads in intervals.icu with the planned loads in the ATP and updates accordingly.

**The fifth script ([5_ATP_PERIOD_NOTE.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/5_ATP_PERIOD_NOTE.py))** makes a note along the whole period (build, transition, race etc.) in intervals.icu.

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
- **Period Notes**: Creates notes for different training periods with color-coded notes based on the period name (e.g., Base, Peak, Race).

## Usage

1. **Setup**: Ensure you have the required Python libraries installed (`logging`, `os`, `pandas`, `requests`).
2. **Configuration**: Update the user variables at the top of the script with your Excel file path, sheet name.
3. **Excel Sheets**: Place the `Intervals_API_Tools_Office365_v1._ATP2intervals_TLA.xlsm` in `C:\TEMP\TLA`. (TLA = e.g. RAA, so rename the sheet to Intervals_API_Tools_Office365_v1._ATP2intervals_[RAA].xlsm).
4. **User_Data**: In the tab User_Data you can put the athlete ID, API key, preferred unitsystem (metric or imperial) and some basic preferences like the color of the note.
5. **Run**: Execute the scripts in the right order to sync your training plan with intervals.icu.
