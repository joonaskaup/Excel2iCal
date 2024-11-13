# From Excel to iCal (`Excel2iCal.py`)

*Note: This code and the readme.md file were created with ChatGPT (version o1-preview or o1-mini). I don’t have prior coding knowledge, and the same bot guided me through navigating GitHub. However, the code appears to work for my purposes.*

## Overview

`Excel2iCal` is a Python script designed to synchronize events from Excel files to your iCal application. It reads event details from specified Excel files and updates the corresponding calendars accordingly.

## Example Image

![Examples](Example_images/example_images.png)

## Features

- **Excel to Calendar Synchronization**: Automatically creates and updates calendar events based on data from Excel files.
- **UID Mapping**: Maintains a unique identifier (UID) for each event to manage updates and deletions effectively.
- **User-Friendly Interface**: Utilizes a Tkinter-based GUI to allow users to select which calendars to synchronize.
- **Configuration Driven**: Easily configure multiple calendars and their corresponding Excel files via a `config.txt` file.

## Limitations

- **Read-Only Synchronization**: 
  - **Excel to Calendar Only**: Changes made directly in the Calendar app **will not** be synchronized back to the Excel files.
  - **Overwrite Risk**: Any modifications or alarms set directly in the Calendar will be overwritten during synchronization.
- **No Conflict Resolution**: If both the Excel file and the Calendar are modified independently, the script does not handle conflict resolution and prioritizes the Excel data.

## Use Case: Enhancing Production Management Workflow

As a Film and TV production manager, managing numerous schedules and coordinating various departments is a complex and time-consuming task. The `Excel2iCal.py` script helps this process by transferring data from Excel spreadsheets to your Calendar application. Here's how it integrates into my daily workflow:

1. **Centralized Data Management**
   - **Challenge**: Handling a vast amount of date and time information.
   - **Solution**: Maintain all relevant dates and times in a single Excel spreadsheet and use the script to transfer this data to multiple calendars, eliminating manual updates and ensuring consistency.

2. **Efficient Calendar Visualization**
   - **Challenge**: Visualizing events across various production calendars is time-consuming.
   - **Solution**: View all events in a consolidated monthly calendar within the Calendar app, making it easier to identify overlaps and plan resources.

3. **Streamlined Department Scheduling**
   - **Example**: Set Department and Location Availability
     - **Process**: Input start and end dates in Excel, transfer to iCal, and share with the team.
     - **Benefit**: Automatic updates ensure the team always has access to the latest information.

4. **Automated Updates and Maintenance**
   - **Challenge**: Manually updating multiple calendars increases the likelihood of errors.
   - **Solution**: Update the Excel spreadsheet and run the script to automatically update all associated calendars.

## Requirements

- Python 3.6 or higher
- Virtual Environment (recommended)

## Installation

1. **Clone the Repository**

   ```bash
   git clone https://github.com/yourusername/your-repo-name.git
   cd your-repo-name

2. **Create a Virtual Environment**

   It is highly recommended to use a virtual environment to manage dependencies.

   ```bash  
   python -m venv venv
   ```

3. **Activate the Virtual Environment**

   ```bash
   source venv/bin/activate
   ```
4. **Install Dependencies**
   
   ```bash
   pip install -r requirements.txt

## Configuration

   Create a config.txt file in the project directory with the following structure:
   ```ini
   [Calendar1]
   Header = Your Header Here
   CalendarName = Your Calendar Name
   ExcelFilePath = /path/to/your/excel1.xlsx
 
   [Calendar2]
   Header = Another Header
   CalendarName = Another Calendar
   ExcelFilePath = /path/to/your/excel2.xlsx

   # Add more calendars as needed
```
- Header: A descriptive header for the configuration section.
- CalendarName: The exact name of the calendar in the macOS Calendar app.
- ExcelFilePath: The full path to the corresponding Excel file containing event data.

## Usage
   Run the script using Python:
   ```bash
   python Excel2iCal.py
   ```

Upon execution:

1. A GUI window will appear listing all configured calendars.
2. Select the calendars you wish to synchronize.
3. Click the "Synchronize" button to start the synchronization process.
**Note**: Ensure that the Excel files are properly formatted with the necessary columns (```Title```,```Start```, ```End```, ```Description```, ```Location```, ```AllDay```).

## Configuration Details

**AllDay Column Configuration**
- Values: The AllDay column should contain boolean values: ```TRUE``` or ```FALSE```.
- ```TRUE```: Marks the event as an all-day event in the Calendar.
- ```FALSE```: The event will include specific start and end times.

   - For events marked as not all-day (```FALSE```), ensure that the ```Start``` and ```End``` columns follow the format ```DD.MM.YYYY HH:MM```. For example:
     - Start: ```25.12.2024 09:00```
     - End: ```25.12.2024 17:00```

- **Example Configuration:**

| Title  | Start  | End  | Description  | Location  | AllDay  |
|:----------|:----------|:----------|:----------|:----------|:----------|
| Christmas Day    | 25.12.2024    | 25.12.2024    | Holiday    | N/A    | TRUE    |
| Team Meeting   | 02.01.2025 10:00    | 02.01.2025 11:00    | Project Update    | Conference Room    | FALSE    |

**Note**: Ensure that the date and time formats are consistent to prevent synchronization issues. Incorrect formatting may lead to events not being created or updated as expected.

**Date and Time Format:**

## Important Notes

- Virtual Environment: Always activate the virtual environment before running the script to ensure that all dependencies are correctly loaded.
- Data Integrity: The script overwrites any changes or alarms made directly in the Calendar app. To maintain data integrity, make all event changes within the Excel files and re-run the synchronization.
- No Two-Way Sync: This tool performs a one-way synchronization from Excel to Calendar. It does not support syncing changes made in the Calendar back to the Excel files.

## Troubleshooting

- Configuration File Missing: Ensure that config.txt exists in the project directory and is correctly formatted.
- Calendar Not Found: Verify that the CalendarName in config.txt matches exactly with the calendar name in the macOS Calendar app.
- Excel File Issues: Ensure that the specified Excel files exist and are accessible. The script will skip any missing or improperly formatted Excel files.

## Future Enhancements

Here are some ideas and features to implement to improve and expand the functionality of `Excel2iCal.py`:

- **Two-Way Synchronization**
  - **Description**: Enable synchronization in both directions, allowing changes made directly in the Calendar app to be reflected back in the Excel files.
  - **Benefits**: This would ensure data consistency across platforms and provide a more flexible and robust synchronization mechanism.

- **Sync Alerts and Notifications**
  - **Description**: Incorporate synchronization of event alerts, reminders, and other notification settings from Excel to Calendar.
  - **Benefits**: Users can manage alert settings within their Excel files, ensuring that notifications are consistently applied in the Calendar app.

- **Automated Email Notifications**
  - **Description**: Implement functionality to automatically send email notifications when changes are detected and synchronized in the Calendar.
  - **Benefits**: Keeps users informed about updates or modifications to their schedules without manually checking the Calendar app.

- **AI-Powered Change Summaries and Reporting**
  - **Description**: Develop an AI agent that utilizes a locally installed language model (LLM) to analyze changes and generate detailed summaries and change reports.
  - **Benefits**: Provides users with clear and concise overviews of what has been updated, enhancing transparency and accountability.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

## License

This project is licensed under the MIT License.


