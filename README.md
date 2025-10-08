# ATP2intervals.icu what does it do?

This is a set of scripts that transfers data from an Annual Training Plan in Excel to intervals.icu. The idea is that you can trickle down a weekly load to different types of sports (activities) and send them to intervals.icu.

The idea is that it must be an easy job to populate intervals.icu with the coach's made ATP. For now, this works quite ok!
I made seperate scripts to keep it maintainable. The first 3 are most important to populate intervals.icu with the data from the ATP. The fourth scripts writes the planned weekly loads per activitytype from the workouts back to the excelsheet. So you can compare the advised weekly load (TSS) with the already planned load. In the excelsheet, you can overwrite the loads (they will be round-off), adjust them and rerun Script 1. Then the planned loads are equal with the target loads.

You can also set time- and distance targets per activity-type, but I kept these away from furter checks. (by now....)

Do you have suggestions for improvement or other ideas? Use this treath on the intervals.icu forum: [ATP2intervals: python scripts to transfer ATP to intervals.icu](https://forum.intervals.icu/t/atp2intervals-python-scripts-to-transfer-atp-to-intervals-icu/91769)

## What's in it

**scripts**

 **The base Python script ([ATP_common_config.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/ATP_common_config.py))** This is the "script" with al (user)variables. To use one common script that is used for input of the others makes updating and managing the scripts easier.

**The first Python script ([1_ATP_LOAD.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/1_ATP_LOAD.py))** automates the process of sending an annual training plan (ATP) with the coach's made ATP to intervals.icu so it wil make weekly targets for TSS (load), time, or distance.
Find more information about ATP on the [intervals.icu forum](https://forum.intervals.icu/t/apps-in-excel-a-guide-to-getting-started/20844).

From:

<img width="1056" height="214" alt="image" src="https://github.com/user-attachments/assets/84a011e8-ed53-4f35-a8c9-fedef38f91da" />

To:

<img width="251" height="234" alt="image" src="https://github.com/user-attachments/assets/41e1b88a-821a-44f9-8e4b-e1ad8536bae0" />




**The second script ([2_ATP_NOTES.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/2_ATP_NOTES.py))** makes notes with multiple comments for your athlete to intervals.icu. It will give information about the main goal, upcoming events and items to focus on. It reads data from the ATP-sheet and sends notes to intervals.icu.

<img width="468" height="196" alt="image" src="https://github.com/user-attachments/assets/ea32cf3f-cb10-4287-8c3d-dedcef4cab49" />


**The third script ([3_ATP_PERIOD_NOTE.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/3_ATP_PERIOD_NOTE.py))** makes a note along the whole period (build, transition, race etc.) in intervals.icu. So the fitness-chart is easier tot read and the different periods are easy to recognise.

In the calendar:

<img width="466" height="203" alt="image" src="https://github.com/user-attachments/assets/d7bd0689-f2ba-4215-ae5c-3c3bb459034f" />

In the Fitness Chart:

<img width="305" height="192" alt="image" src="https://github.com/user-attachments/assets/af0fbb1d-68c8-4063-a279-eb58fd992364" />


**The fourth script ([4_LOAD_CHECK.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/4_LOAD_CHECK.py))** compares the planned target loads in intervals.icu with the planned loads in the ATP and updates accordingly.

**The fifth script ([5_ATP_WEEKLY_LOAD_FEEDBACK_NOTES.py](https://github.com/richardadvocaat/ATP2intervals.icu/blob/main/5_ATP_WEEKLY_LOAD_FEEDBACK_NOTES.py))** checks if the athlete was compliant to the plan based on the reported load, time, and distance in intervals.icu. This script can be used to automate the proces to check if the athlete was compliant to the ATP. Not that usefull for a truely committed coaches or trainera, because normaly this is not something that can be automated. ;-)

Be compliant or....

<img width="468" height="204" alt="image" src="https://github.com/user-attachments/assets/3bcc4ecc-b93d-49a8-9b96-8ac985b79358" />



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

1. **Setup**: Ensure you have the required Python libraries installed (`logging`, `os`, `pandas`, `requests`, `openpyxl`).
2. **Configuration**: Update the user variables at the top of the script with your Excel file path, sheet name.
3. **Excel Sheets**: Place the `ATP2intervals_TLA_YYYY.xlsm` in `C:\TEMP\TLA`. (TLA = e.g. RAA and YYYY = 2026, so rename the sheet to ATP2intervals_RAA_2026.xlsm).
4. **User_Data**: In the tab User_Data you can put the athlete ID, API key, preferred unit system (metric or imperial) and some basic preferences like the color of the note.
5. **Set up ATP**: Fill in the race calendar and some basic stuff like start and end of the ATP-period, yearly TSS. Fill in the ATP period per week (e.g. BASE 1-1) and the recommended load is listed in the next collumn. Then you can distribute the load to different sport types like RUN, RIDE and SWIM. In the sheet there is a check if the total of different weekly loads are equal to the recommended load.
7. **Run**: Execute the scripts in the right order to sync your training plan with intervals.icu.
8. **Check**: After running the first scripts the target_loads are placed in intervals.icu. 4_LOAD_CHECK.py rerieves the actual planned load from the trainingplans. So you can easily check if plans in intervals.icu are compliant with the ATP.
9. **Check_Athlete**: Run 5_ATP_WEEKLY_LOAD_FEEDBACK_NOTES to check if the athlete ws compliant to the given load. (This is is that serious..)

## To Do

1. Store parameters for the coach in a separate document where the coached athlete names are also stored. (or get this list from intervals.icu.)
