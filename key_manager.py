import os
import base64
from firebase_client import db

class KeyManager:
    """
    Manages encryption key upload and retrieval using Firestore.
    """
    def __init__(self, key_file="encryption_key.key"):
        self.key_file = key_file
        if db is None:
            print("Warning: Firestore client is not initialized.")

    def upload_key(self):
        """
        Uploads the encryption key to Firestore (config/secrets).
        """
        if db is None:
            print("Error: Firestore client is not initialized.")
            return

        if not os.path.exists(self.key_file):
            print(f"Error: Key file '{self.key_file}' does not exist.")
            return

        try:
            with open(self.key_file, "rb") as f:
                key_bytes = f.read()
                key_b64 = base64.b64encode(key_bytes).decode('utf-8')

            doc_ref = db.collection('config').document('secrets')
            doc_ref.set({
                'encryption_key': key_b64
            }, merge=True)
            
            print(f"Key uploaded to Firestore (config/secrets).")
            
        except Exception as e:
            print(f"Error uploading key: {e}")

    def retrieve_key(self, download_path=None):
        """
        Retrieves the encryption key from Firestore.
        """
        if db is None:
            print("Error: Firestore client is not initialized.")
            return

        if download_path is None:
            download_path = self.key_file

        try:
            doc_ref = db.collection('config').document('secrets')
            doc = doc_ref.get()
            
            if not doc.exists:
                print("Error: Secrets document not found in Firestore.")
                return

            data = doc.to_dict()
            key_b64 = data.get('encryption_key')
            
            if not key_b64:
                print("Error: Encryption key not found in secrets document.")
                return

            key_bytes = base64.b64decode(key_b64)
            
            with open(download_path, "wb") as f:
                f.write(key_bytes)
                
            print(f"Key downloaded to {download_path}.")
            
        except Exception as e:
            print(f"Error retrieving key: {e}")

if __name__ == "__main__":
    manager = KeyManager()
    # manager.upload_key() # Uncomment to upload
    # manager.retrieve_key("downloaded_key.key") # Uncomment to test download
