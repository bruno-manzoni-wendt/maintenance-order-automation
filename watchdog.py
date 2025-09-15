# File Watchdog Script
# Monitors an Excel file for changes and triggers automated processing

import subprocess
from datetime import datetime, time
from time import sleep
import os

# File paths configuration
base_path = r'path\to\project\data'
file_name = r'WorkOrders.xlsm'
file_path = os.path.join(base_path, file_name)

timestamp_file_path = r'path\to\project\scripts\timestamp.txt'

# Read last modification timestamp
with open(timestamp_file_path, 'r') as txt:
    last_modified = float(txt.read())

print(f"\n======= WATCHING {file_name} =======")

while True:
    if os.path.exists(file_path):  # Check if file exists in directory
        if last_modified != os.stat(file_path).st_mtime:  # If file was modified:
            
            print(f'\n{datetime.now().time().strftime("%H:%M")} {file_name} UPDATED\n')
            last_modified = os.stat(file_path).st_mtime  # Update variable with new timestamp
            
            # Write new timestamp to file
            with open(timestamp_file_path, 'w') as txt:
                txt.write(str(last_modified))
            
            # Execute main processing script
            subprocess.run(["python", r"path\to\project\scripts\process_work_orders.py"])
            
            print(f"\n======= WATCHING {file_name} =======")

    # Stop monitoring after specified time
    if datetime.now().time() > time(16, 50):
        for i in range(5, 0, -1):
            print(f'\nClosing in {i} seconds', end='\r', flush=True)
            sleep(1)

    sleep(20)  # Check for changes every 20 seconds