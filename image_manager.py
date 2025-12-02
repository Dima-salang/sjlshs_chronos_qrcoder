import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
import pickle

class DriveImageManager:
    """
    Manages image uploads to Google Drive.
    """
    def __init__(self, images_dir="images", credentials_file="credentials.json", token_file="token.pickle"):
        self.images_dir = images_dir
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.service = None
        self.folder_cache = {} # Cache folder IDs to avoid repeated API calls

    def authenticate(self):
        """
        Authenticates with Google Drive API.
        Prioritizes Service Account if credentials.json is a service account key.
        Otherwise falls back to OAuth flow (simplified here).
        """
        try:
            # Check if we have a service account file
            if os.path.exists(self.credentials_file):
                # Assuming service account for simplicity in automation
                # For OAuth user consent, we would use InstalledAppFlow
                self.creds = service_account.Credentials.from_service_account_file(
                    self.credentials_file, scopes=['https://www.googleapis.com/auth/drive']
                )
                self.service = build('drive', 'v3', credentials=self.creds)
                print("Authenticated with Google Drive (Service Account).")
            else:
                print(f"Warning: Credentials file '{self.credentials_file}' not found.")
        except Exception as e:
            print(f"Authentication failed: {e}")

    def get_folder_id(self, folder_name, parent_id=None):
        """
        Finds or creates a folder in Drive.
        """
        if not self.service:
            return None

        cache_key = f"{parent_id}_{folder_name}"
        if cache_key in self.folder_cache:
            return self.folder_cache[cache_key]

        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        if parent_id:
            query += f" and '{parent_id}' in parents"
        
        results = self.service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])

        if files:
            folder_id = files[0]['id']
        else:
            # Create folder
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            if parent_id:
                file_metadata['parents'] = [parent_id]
            
            folder = self.service.files().create(body=file_metadata, fields='id').execute()
            folder_id = folder.get('id')
            print(f"Created folder '{folder_name}' (ID: {folder_id})")

        self.folder_cache[cache_key] = folder_id
        return folder_id

    def upload_images(self):
        """
        Uploads images to Google Drive, maintaining folder structure.
        """
        if not self.service:
            self.authenticate()
            if not self.service:
                print("Error: Drive service not initialized.")
                return

        if not os.path.exists(self.images_dir):
            print(f"Error: Images directory '{self.images_dir}' does not exist.")
            return

        print(f"Scanning '{self.images_dir}' for images...")
        
        # Root folder for uploads
        root_folder_name = "Chronos_Images"
        root_id = self.get_folder_id(root_folder_name)

        for root, dirs, files in os.walk(self.images_dir):
            for file in files:
                if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                    local_path = os.path.join(root, file)
                    
                    # Determine parent folder in Drive based on local structure
                    # e.g. images/SectionA/student.jpg -> Drive/Chronos_Images/SectionA/student.jpg
                    
                    relative_path = os.path.relpath(root, self.images_dir)
                    current_parent_id = root_id
                    
                    if relative_path != ".":
                        # Create/Find subfolders
                        parts = relative_path.split(os.sep)
                        for part in parts:
                            current_parent_id = self.get_folder_id(part, current_parent_id)
                    
                    # Check if file exists
                    query = f"name = '{file}' and '{current_parent_id}' in parents and trashed = false"
                    results = self.service.files().list(q=query, fields="files(id)").execute()
                    if results.get('files'):
                        print(f"Skipping {file} (already exists).")
                        continue

                    # Upload file
                    file_metadata = {'name': file, 'parents': [current_parent_id]}
                    media = MediaFileUpload(local_path, resumable=True)
                    
                    print(f"Uploading {local_path}...")
                    self.service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                    print(f"Uploaded {file}.")

        print("Image upload complete.")

if __name__ == "__main__":
    manager = DriveImageManager()
    manager.upload_images()
