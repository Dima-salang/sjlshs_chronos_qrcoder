import firebase_admin
from firebase_admin import credentials, firestore
import sys
import os

db = None
try:
    # IMPORTANT: Replace with the actual path to your key file
    key_path = "sjlshs-chronos-firebase-adminsdk-fbsvc-82e3ea3498.json" 
    
    if not os.path.exists(key_path):
        raise FileNotFoundError(f"Firebase service account key not found at: {key_path}")

    cred = credentials.Certificate(key_path)
    firebase_admin.initialize_app(cred)
    db = firestore.client()
    print("Successfully connected to Firestore.")

except Exception as e:
    print(f"Failed to initialize Firestore: {e}", file=sys.stderr)
    # In a real app, you might want to show a pop-up error message to the user
    # and then exit, as the app cannot function without the database.
    db = None # Ensure db is None if initialization fails
