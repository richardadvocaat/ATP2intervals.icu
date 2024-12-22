# ATP2intervals.icu

This Python script automates the process of sending an annual training plan with weekly loads (for multiple sports) to intervals.icu. It reads data from an Excel file and updates or creates events accordingly. The tab in the Excel file used for this script can be integrated into the sheets for planning an Annual Training Plan that can be found here: https://drive.google.com/drive/folders/1WhIOf2XkGiZBEN_7tX2PSShmF-QXBnBF


## Features

- **Read Excel File**: The script reads an Excel file containing weekly load, time and distance targets for different activities (Bike, Run, Swim).
- **Create or Update Events**: It creates new events or updates existing ones in intervals.icu based on the load targets.
- **Delete Events**: If the load target is zero, the script deletes the corresponding event in intervals.icu.
- **Efficient Data Handling**: The script retrieves all relevant events in a single API call to improve performance and reduce server load.
- **Unit Conversion**: The script allows users to specify their unit preference (metric or imperial) for Bike and Run distances, while Swim distances remain in meters.
- **Custom Descriptions**: The script now supports adding custom descriptions based on the 'period' and 'focus' columns from the Excel file.
  - If 'period' is "Rest", the description will include "Stay in bed or on the beach!".
  - If 'focus' has a value, the description will include "Focus this week on {focus}".

## Usage

1. **Setup**: Ensure you have the required Python libraries installed (`pandas`, `requests`).
2. **Configuration**: Update the user variables at the top of the script with your Excel file path, sheet name, athlete ID, API key, and preferred unit system (metric or imperial).
3. **Run**: Execute the script to sync your training plan with intervals.icu.
