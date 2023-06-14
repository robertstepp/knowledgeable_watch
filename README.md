# Knowledgeable_Watch

This PowerShell script allows users to track time entries for various items based on a configuration file. The script
provides a graphical user interface (GUI) where users can select multiple items and track their time. It generates a 
timesheet report in an Excel workbook, including the total duration for each item.

## Features

- Reads a configuration file (`config.ini`) that contains a list of items with their corresponding categories.
- Presents a GUI with buttons for each item, allowing the user to select multiple items.
- Calculates the total duration for the selected items and generates a timesheet report.
- Saves the timesheet report as an Excel workbook (`Timesheet-YYYYMMDD.xlsx`) with the current date.
- Timesheet report includes the total duration for each item, rounded to the nearest minute.
- The timesheet report also expresses the duration in tenths per hour, rounded up to the first digit.
- Checks if an Excel spreadsheet for the current date already exists and prompts for confirmation before overwriting.
- Creates a raw timesheet CSV file (`raw_timesheet.csv`) to store the start and end times of each item selection.
- Includes the raw timesheet data as a separate sheet in the generated Excel workbook.

## Usage

1. Ensure that a configuration file named `config.ini` is present in the same directory as the script.
2. Run the script to launch the GUI.
3. Select the desired items by clicking the corresponding buttons.
4. Click the "Done" button to generate the timesheet report.
5. Confirm overwriting an existing timesheet, if applicable.

## Configuration File Format (config.ini)

The `config.ini` file should be formatted as follows:

```plaintext
[Category1]
Item1
Item2
...

[Category2]
Item3
Item4
...

## Debugging

You can use the Write-Debug cmdlet to output debug information. Set `$DebugPreference = 'Continue'` at the top of the script to enable debug output.
### Contact

If you have any questions or issues, please contact Robert Stepp at robert@robertstepp.ninja.