# Automated Maintenance Request Processing System

**Situation:** Maintenance Service Orders were submitted manually on paper, making it difficult to track, monitor, and analyze requests.

**Task:** Replace the manual process with a digital solution that enables traceability and monitoring.

**Action:** Designed Excel-based forms integrated with Python automation, which register the requests and generated standardized Service Order documents automatically. These documents were sent to be printed in both the Manager’s and Maintenance offices.

**Result:** Fully digitized the request process, reducing processing time and enabling monitoring and reporting of maintenance requests.


## How It Works

### File Watcher Script (watchdog.py)

- Monitors Excel file for changes with timestamp tracking
- A VBA script register it and saves the file when a new requests is created on Excel
- When the modified timestammp of the Excel file changes, that will automatically trigger the 'service_order_processor.py' script

###  Maintenance Request Processor (service_order_processor.py)

- Reads maintenance requests from Excel spreadsheet
- Generates formatted Word documents from template
- Inserts digital signatures and request data automatically
- Prints documents to multiple network printers
- Sends authorization emails for improvement requests via Outlook
- Tracks processed requests to avoid duplicates
- Includes comprehensive error handling and logging

### Key Features:

- File monitoring with modification timestamp detection
- Automated document generation using python-docx
- Multi-printer support
- Digital signature insertion from image files
- Text formatting and wrapping for document templates
- Excel data validation and processing
