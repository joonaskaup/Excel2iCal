import pandas as pd
from appscript import app, k
from datetime import datetime
import traceback
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import configparser
from pathlib import Path

def sync_excel_to_calendar():
    print("Entered sync_excel_to_calendar()")
    try:
        # Read configurations from the config file
        config_file_path = 'config.txt'
        if not os.path.exists(config_file_path):
            print(f"Configuration file '{config_file_path}' not found.")
            return

        configs = read_config_file(config_file_path)
        if not configs:
            print("No configurations found in the config file.")
            return

        # Load last synchronization times
        sync_times_file = 'sync_times.json'
        if os.path.exists(sync_times_file):
            with open(sync_times_file, 'r') as f:
                sync_times = json.load(f)
        else:
            sync_times = {}

        # Display configurations and synchronization status
        selected_configs = select_configs(configs, sync_times)

        if not selected_configs:
            print("No calendars selected for synchronization. Exiting.")
            return

        # Access the Calendar app
        print("Accessing Calendar app...")
        calendar_app = app('Calendar')
        print("Calendar app accessed.")

        for config in selected_configs:
            calendar_name = config['CalendarName']
            excel_file_path = config['ExcelFilePath']
            print(f"\nProcessing Calendar: {calendar_name}")
            print(f"Excel File: {excel_file_path}")

            if not os.path.exists(excel_file_path):
                print(f"Excel file '{excel_file_path}' not found. Skipping.")
                continue

            # Read Excel data
            print("Reading Excel file...")
            df = pd.read_excel(excel_file_path)
            print("Excel file read successfully.")

            if df.empty:
                print("DataFrame is empty. Skipping.")
                continue

            # Remove rows where all elements are NaN
            df.dropna(how='all', inplace=True)

            # Update last modified time of Excel file
            excel_modified_time = os.path.getmtime(excel_file_path)
            excel_modified_datetime = datetime.fromtimestamp(excel_modified_time)

            # Load UID mapping
            uid_mapping_file = f'uid_mapping_{calendar_name}.json'
            if os.path.exists(uid_mapping_file):
                with open(uid_mapping_file, 'r') as f:
                    uid_mapping = json.load(f)
                print("UID mapping loaded.")
            else:
                uid_mapping = {}
                print("No UID mapping file found. Starting with an empty mapping.")

            # Select the target calendar
            try:
                target_calendar = calendar_app.calendars[calendar_name]
            except Exception as e:
                print(f"Calendar '{calendar_name}' not found. Creating it.")
                # Create the calendar if it doesn't exist
                target_calendar = calendar_app.make(new=k.calendar, with_properties={k.name: calendar_name})

            # Get existing events in the target calendar
            print("Retrieving existing events...")
            existing_events = {evt.uid(): evt for evt in target_calendar.events()}
            print(f"Found {len(existing_events)} existing events.")

            # Events processed in this run
            processed_event_uids = set()

            # Collect all event_keys present in Excel
            current_event_keys = set()

            # Iterate over events in Excel
            print("Processing events from Excel...")
            for index, row in df.iterrows():
                # Check if the row is empty or missing required fields
                if pd.isnull(row.get('Title')) and pd.isnull(row.get('Start')) and pd.isnull(row.get('End')):
                    print(f"Row {index} is empty or missing required fields. Skipping.")
                    continue  # Skip to the next row

                title = row.get('Title', 'Untitled Event')
                start = row.get('Start')
                end = row.get('End')
                description = row.get('Description', '')
                location = row.get('Location', '')
                all_day = row.get('AllDay', False)

                # Check if essential fields are missing
                if pd.isnull(title) or pd.isnull(start) or pd.isnull(end):
                    print(f"Row {index} is missing required fields. Skipping.")
                    continue  # Skip to the next row

                # Replace NaN with empty string for description and location
                if pd.isnull(description):
                    description = ''
                if pd.isnull(location):
                    location = ''

                # Ensure 'all_day' is a boolean
                if pd.isnull(all_day):
                    all_day = False
                else:
                    if isinstance(all_day, str):
                        all_day_str = all_day.strip().lower()
                        if all_day_str == 'true':
                            all_day = True
                        elif all_day_str == 'false':
                            all_day = False
                        else:
                            all_day = False
                    else:
                        all_day = bool(all_day)

                # Generate a unique event key based on original Excel data
                try:
                    original_start = pd.to_datetime(start, dayfirst=True)
                    original_end = pd.to_datetime(end, dayfirst=True)
                    event_key = f"{title}_{original_start.isoformat()}_{original_end.isoformat()}"
                    print(f"\nEvent Key: {event_key}")
                    print(f"Original Start: {original_start.isoformat()}, Original End: {original_end.isoformat()}")
                except Exception as e:
                    print(f"Date parsing error for event '{title}': {e}")
                    traceback.print_exc()
                    continue  # Skip to the next event

                current_event_keys.add(event_key)

                # Handle both all-day and timed events
                try:
                    if all_day:
                        start_date = original_start.date()
                        end_date = original_end.date()
                        print(f"All-day event: Start Date: {start_date}, End Date: {end_date}")

                        start_datetime = datetime.combine(start_date, datetime.min.time())
                        end_datetime = datetime.combine(end_date, datetime.max.time())
                    else:
                        # Timed event
                        start_datetime = original_start.to_pydatetime().replace(tzinfo=None)
                        end_datetime = original_end.to_pydatetime().replace(tzinfo=None)
                        print(f"Timed Event: Start: {start_datetime}, End: {end_datetime}")

                except Exception as e:
                    print(f"Date adjustment error for event '{title}': {e}")
                    traceback.print_exc()
                    continue  # Skip to the next event

                # Check if event already exists in UID mapping
                mapping_entry = uid_mapping.get(event_key)
                uid = mapping_entry.get('uid') if mapping_entry else None

                if uid and uid in existing_events:
                    # Update existing event
                    evt = existing_events[uid]

                    # Update event properties
                    try:
                        evt.summary.set(title)
                        evt.start_date.set(start_datetime)
                        evt.end_date.set(end_datetime)
                        evt.description.set(description)
                        evt.location.set(location)
                        evt.allday_event.set(all_day)
                        print(f"Updated event: {title} (UID: {uid})")
                    except Exception as e:
                        print(f"Error updating event '{title}': {e}")
                        traceback.print_exc()
                        continue  # Skip to the next event
                else:
                    # Create new event
                    event_properties = {
                        k.summary: title,
                        k.start_date: start_datetime,
                        k.end_date: end_datetime,
                        k.description: description,
                        k.location: location,
                        k.allday_event: all_day
                    }
                    try:
                        new_event = target_calendar.make(new=k.event, with_properties=event_properties)
                        new_uid = new_event.uid()

                        # Save UID and original start and end for this event
                        uid_mapping[event_key] = {
                            'uid': new_uid,
                            'original_start': original_start.isoformat(),
                            'original_end': original_end.isoformat()
                        }
                        print(f"Created new event: {title} (UID: {new_uid})")
                    except Exception as e:
                        print(f"Error creating event '{title}': {e}")
                        traceback.print_exc()
                        continue  # Skip to the next event

                # Add UID to processed events set
                if uid:
                    processed_event_uids.add(uid)
                else:
                    processed_event_uids.add(uid_mapping[event_key]['uid'])

            # Identify event_keys that have been removed from Excel
            mapped_event_keys = set(uid_mapping.keys())
            deleted_event_keys = mapped_event_keys - current_event_keys

            if deleted_event_keys:
                print(f"\nFound {len(deleted_event_keys)} deleted event(s). Deleting from Calendar...")
                for event_key in deleted_event_keys:
                    uid = uid_mapping[event_key]['uid']
                    if uid in existing_events:
                        try:
                            existing_events[uid].delete()
                            print(f"Deleted event '{event_key}' with UID: {uid}")
                        except Exception as e:
                            print(f"Error deleting event '{event_key}' (UID: {uid}): {e}")
                            traceback.print_exc()
                    del uid_mapping[event_key]
            else:
                print("\nNo deleted events found.")

            # Save the updated UID mapping
            try:
                with open(uid_mapping_file, 'w') as f:
                    json.dump(uid_mapping, f, indent=4)
                print("UID mapping saved.")
            except Exception as e:
                print(f"Error saving UID mapping: {e}")
                traceback.print_exc()

            # Update synchronization time
            sync_times[calendar_name] = datetime.now().isoformat()

        # Save synchronization times
        with open(sync_times_file, 'w') as f:
            json.dump(sync_times, f, indent=4)

        print("\nSynchronization complete.")

    except Exception as e:
        print(f"An error occurred in sync_excel_to_calendar(): {e}")
        traceback.print_exc()

def read_config_file(config_file_path):
    config = configparser.ConfigParser()
    config.optionxform = str  # Preserve case sensitivity
    config.read(config_file_path)
    configs = []
    for section in config.sections():
        header = config.get(section, 'Header', fallback='No Header')
        calendar_name = config.get(section, 'CalendarName')
        excel_file_path = config.get(section, 'ExcelFilePath')
        configs.append({
            'Header': header,
            'CalendarName': calendar_name,
            'ExcelFilePath': excel_file_path
        })
    return configs

def select_configs(configs, sync_times):
    """
    Display the configurations with their synchronization status and allow the user to select which ones to process.
    """
    selection_window = tk.Tk()
    selection_window.title("Select Calendars to Synchronize")
    selection_window.geometry("800x600")
    selection_window.resizable(False, False)

    selected_configs = []

    def on_submit():
        for var, config in checkbox_vars:
            if var.get():
                selected_configs.append(config)
        if selected_configs:
            selection_window.destroy()
        else:
            messagebox.showwarning("No Selection", "Please select at least one calendar to synchronize.")

    label = tk.Label(selection_window, text="Select calendars to synchronize:", font=("Helvetica", 12))
    label.pack(pady=10)

    frame = tk.Frame(selection_window)
    frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    canvas = tk.Canvas(frame)
    scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    checkbox_vars = []
    for config in configs:
        header = config['Header']
        calendar_name = config['CalendarName']
        excel_file_path = config['ExcelFilePath']
        excel_modified_time = os.path.getmtime(excel_file_path) if os.path.exists(excel_file_path) else None
        excel_modified_datetime = datetime.fromtimestamp(excel_modified_time).strftime('%Y-%m-%d %H:%M:%S') if excel_modified_time else "File Not Found"
        last_sync_time = sync_times.get(calendar_name, "Never")

        up_to_date = False
        if excel_modified_time and last_sync_time != "Never":
            last_sync_timestamp = datetime.fromisoformat(last_sync_time).timestamp()
            up_to_date = excel_modified_time <= last_sync_timestamp

        status_text = "Up to date" if up_to_date else "Needs sync"
        status_color = "green" if up_to_date else "red"

        var = tk.IntVar()
        cb = tk.Checkbutton(scrollable_frame, text=header, variable=var)
        cb.pack(anchor='w')

        info_text = f"Calendar Name: {calendar_name}\nExcel File: {excel_file_path}\nExcel Last Modified: {excel_modified_datetime}\nLast Sync: {last_sync_time}\nStatus: {status_text}"
        info_label = tk.Label(scrollable_frame, text=info_text, fg=status_color, justify='left', wraplength=700)
        info_label.pack(anchor='w', padx=20, pady=5)

        checkbox_vars.append((var, config))

    submit_button = tk.Button(selection_window, text="Synchronize", command=on_submit)
    submit_button.pack(pady=10)

    selection_window.mainloop()

    return selected_configs

if __name__ == '__main__':
    try:
        print("Script started.")
        sync_excel_to_calendar()
    except Exception as e:
        print(f"An unhandled exception occurred: {e}")
        traceback.print_exc()
