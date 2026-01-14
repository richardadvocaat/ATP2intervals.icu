# ATP2intervals.icu — What does it do?

These scripts transfer data from an Annual Training Plan (ATP) in Excel to intervals.icu. They distribute weekly training loads across different activity types and upload those targets to intervals.icu.

The idea is to make it easy to populate intervals.icu with a coach-created ATP. Currently, this works reasonably well. The scripts are separated to keep the codebase maintainable; the first three scripts are the most important for populating intervals.icu with ATP data. The fourth script writes planned weekly loads per activity type back to the Excel workbook so you can compare advised weekly load (TSS) with planned load.

## What's included

**Scripts**

- **ATP_common_config.py** — Configuration and shared variables.
- **1_ATP_LOAD.py** — Sends an annual training plan (ATP) to intervals.icu and creates weekly targets for TSS (load), time, or distance.
  
From:

<img width="1056" height="214" alt="image" src="https://github.com/user-attachments/assets/84a011e8-ed53-4f35-a8c9-fedef38f91da" />

To:

<img width="251" height="234" alt="image" src="https://github.com/user-attachments/assets/41e1b88a-821a-44f9-8e4b-e1ad8536bae0" />

- **2_ATP_NOTES.py** — Adds period descriptions and notes for ATP periods on intervals.icu.
- **3_ATP_PERIOD_NOTE.py** — Creates a note covering an entire ATP period (build, transition, race, etc.) for clearer interpretation of the fitness chart.
In the calendar:

<img width="466" height="203" alt="image" src="https://github.com/user-attachments/assets/d7bd0689-f2ba-4215-ae5c-3c3bb459034f" />

In the Fitness Chart:

<img width="305" height="192" alt="image" src="https://github.com/user-attachments/assets/af0fbb1d-68c8-4063-a279-eb58fd992364" />

- **4_LOAD_CHECK.py** — Compares planned target loads in intervals.icu with the ATP and updates the workbook where needed.
- **5_ATP_WEEKLY_LOAD_FEEDBACK_NOTES.py** — Evaluates weekly compliance with the ATP and optionally creates feedback notes.

<img width="468" height="204" alt="image" src="https://github.com/user-attachments/assets/3bcc4ecc-b93d-49a8-9b96-8ac985b79358" />
  
- **6_RACES.py** — Exports race events to the workbook.
- **NOTE_REMOVER.py** — Removes NOTE events matching a specific year and keyword.

## Features

- Read Excel file containing weekly load, time, and distance targets for activities (Bike, Run, Swim).
- Create or update events on intervals.icu based on load targets.
- Delete events when the load target is zero.
- Efficient data retrieval via single API calls where possible.
- Unit conversion options (metric or imperial) for Bike and Run distances; Swim distances remain in meters.
- Adds comments for tests specified in the 'test' column and adds focus based on specified focus columns.
- Adds custom race category descriptions and personalized messages using the athlete's name from intervals.icu.

## Usage

1. Install required Python libraries (`logging`, `os`, `pandas`, `requests`, `openpyxl`, `xlwings`).
2. Update the user variables in ATP_common_config.py (Excel path, sheet names, API keys).
3. Place the ATP2intervals_TLA_YYYY.xlsm file in `C:\TEMP\TLA`. (TLA—for example, RAA; YYYY—for example, 2026). Rename the file accordingly (e.g., `ATP2intervals_RAA_2026.xlsm`).
4. In the workbook tab `User_Data`, provide the athlete ID, API key, preferred unit system (metric or imperial), and basic preferences such as note color.
5. Fill in the race calendar and ATP period data. The recommended load is listed in the next column. You can then distribute the recommended load to activity types such as RUN, RIDE, and SWIM.
6. Run the scripts in the proper order to sync your ATP with intervals.icu.
7. After the initial sync, run `4_LOAD_CHECK.py` to retrieve the planned loads from intervals.icu and compare them with the ATP. Use `5_ATP_WEEKLY_LOAD_FEEDBACK_NOTES.py` to generate feedback notes about compliance (use thoughtfully — this is intended as a light, automated check rather than a definitive evaluation).

## To Do

1. Store coach-specific parameters and athlete lists in a separate configuration.
2. Centralize variables currently spread across scripts and workbook sheets so all configuration is in a single location.
