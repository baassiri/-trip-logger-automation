from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import os
from config import FILE_PATH  # Ensure this points to the correct `data` directory

def authenticate_drive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()  # Authenticate via OAuth
    return GoogleDrive(gauth)

def clear_google_drive_xlsm(file_id):
    drive = authenticate_drive()
    file = drive.CreateFile({'id': file_id})
    
    try:
        print("🚀 Overwriting file with a blank version...")
        open(FILE_PATH, 'w').close()  # Clears local file
        file.SetContentFile(FILE_PATH)
        file.Upload()
        print("✅ File cleared and uploaded successfully.")
    except Exception as e:
        print(f"❌ Failed to clear file: {e}")

if __name__ == "__main__":
    FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"  # Update with your file ID
    clear_google_drive_xlsm(FILE_ID)
