# ATP2intervals.icu

This Python script automates the process of sending an annual training plan with weekly loads (for Bike, Run, and Swim) to intervals.icu. It reads data from an Excel file and updates or creates events in intervals.icu based on the provided load targets.

## Features

- **Read Excel File**: The script reads an Excel file containing weekly load targets for different activities (Bike, Run, Swim).
- **Create or Update Events**: It creates new events or updates existing ones in intervals.icu based on the load targets.
- **Delete Events**: If the load target is zero, the script deletes the corresponding event in intervals.icu.
- **Efficient Data Handling**: The script retrieves all relevant events in a single API call to improve performance and reduce server load.

## Usage

1. **Setup**: Ensure you have the required Python libraries installed (`pandas`, `requests`).
2. **Configuration**: Update the user variables at the top of the script with your Excel file path, sheet name, athlete ID, and API key.
3. **Run**: Execute the script to sync your training plan with intervals.icu.

It focusses on load targets now, but apply time or distance targets can be added if needed.
