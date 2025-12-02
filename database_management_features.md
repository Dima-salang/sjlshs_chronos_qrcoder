# Database Management Features Added

## Summary
Added functionality to upload local database to Firebase and delete records from both local and Firestore databases.

## Changes Made

### 1. data_importer.py
Added `MasterListManager` class with the following methods:
- `get_local_records()`: Retrieves all records from the local SQLite database
- `upload_local_to_firestore(progress_callback=None)`: Uploads all local records to Firestore with batch processing
- `delete_local_records()`: Deletes all records from the local SQLite database
- `delete_firestore_records(progress_callback=None)`: Deletes all records from Firestore with batch processing

### 2. main.py
#### UI Changes:
- Added "Database Management" section in the Master List Import tab
- Added three new buttons:
  - **Upload Local DB to Firebase**: Uploads all local database records to Firestore
  - **Delete Local DB**: Deletes all records from the local database
  - **Delete Firestore DB**: Deletes all records from Firestore

#### Handler Methods:
- `_upload_local_to_firebase()`: Confirms and starts upload process
- `_run_upload_local_to_firebase()`: Executes upload in separate thread
- `_delete_local_db()`: Confirms and starts local deletion
- `_run_delete_local_db()`: Executes local deletion in separate thread
- `_delete_firestore_db()`: Confirms and starts Firestore deletion
- `_run_delete_firestore_db()`: Executes Firestore deletion in separate thread

## Features
- **Batch Processing**: Both upload and delete operations use batch processing (400 records per batch) for efficiency
- **Progress Callbacks**: Operations log progress to the UI
- **Confirmation Dialogs**: All destructive operations require user confirmation
- **Thread Safety**: All long-running operations execute in separate threads to prevent UI freezing
- **Error Handling**: Comprehensive error handling with user-friendly error messages

## Usage
1. **Upload Local DB to Firebase**: Click the button to sync all local records to Firestore
2. **Delete Local DB**: Click to permanently delete all local records (with confirmation)
3. **Delete Firestore DB**: Click to permanently delete all Firestore records (with confirmation)

All operations log their progress in the Import Log text area.
