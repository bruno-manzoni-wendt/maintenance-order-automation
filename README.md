# maintenance-order-automation
Python automation system with Excel to register and generate Maintenance Service Order requests documents to be printed.


# Automated Maintenance Request Processing System

## How It Works - Step by Step

### 1. **File Monitoring Phase** (Watchdog Script)
- Script continuously monitors a specific Excel file (`maintenance_requests.xlsm`) for changes
- Uses file modification timestamps to detect when the file has been updated
- Runs in a loop every 20 seconds, checking if the file's last modified time has changed
- Operates during business hours only (closes at 4:50 PM weekdays, 3:50 PM Fridays)

### 2. **Change Detection & Trigger**
- When the Excel file is modified (new maintenance request added):
  - Updates the stored timestamp
  - Automatically launches the maintenance request processor script
  - Continues monitoring for additional changes

### 3. **Data Processing Phase** (Processor Script)
- Reads the Excel file and compares work order numbers
- Identifies new requests by comparing the latest Excel entry with a tracking file
- If no new requests exist, the script exits gracefully

### 4. **Document Generation**
For each new maintenance request:
- **Load Template**: Opens a Word document template
- **Extract Data**: Pulls request details (requester name, department, description, dates, etc.)
- **Populate Document**: 
  - Adds work order number to header
  - Fills in requester information and service details
  - Formats long descriptions with proper line wrapping
- **Insert Signatures**: Automatically adds digital signature images for requester and supervisor

### 5. **Document Output**
- **Save**: Creates a new Word document with a unique filename
- **Print**: Sends document to multiple network printers simultaneously
  - Regular printers get all documents
  - Color printer used sparingly (5% of time) to save costs
- **Track**: Updates tracking file to prevent reprocessing the same request

### 6. **Authorization Workflow** (For Improvement Requests)
- If the request type is "IMPROVEMENT":
  - Composes an authorization email with request details
  - Sends via Outlook to managers for approval
  - Uses GUI automation to complete the send process

### 7. **Error Handling & Logging**
- Captures and logs printing errors to a backlog file
- Validates data completeness before processing
- Handles COM interface connections safely
- Provides user feedback throughout the process

### 8. **Cleanup & Shutdown**
- At end of business day, runs cleanup scripts
- Closes all connections properly
- Provides countdown notification before shutdown

## Key Benefits
- **Zero Manual Intervention**: Fully automated from detection to printing
- **Real-time Processing**: Requests are processed within minutes of submission
- **Multi-format Output**: Generates both digital and physical copies
- **Approval Integration**: Automatically routes improvement requests for authorization
- **Cost Optimization**: Smart printer selection reduces operational costs
- **Error Recovery**: Comprehensive logging ensures no requests are lost
