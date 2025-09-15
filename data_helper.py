import os
import io
import pickle
import platform
import subprocess
import app_context
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import json
from window_helper import ToplevelWindowHelper
import tkinter as tk
from tkinter import messagebox
import sys
import time

class DriveManager:
    """
    A helper class to manage Google Drive operations such as authentication, file download, and file upload.

    Parameters
    ----------
    credentials_file (str, optional): Default = 'credentials.json' 
        Path to the JSON credentials file used for OAuth 2.0 authentication with Google Drive.    
    token_file (str, optional): Default = 'token.json'
        Path to the local JSON token file that caches access credentials to avoid re-authentication 
        on each run.    

    Attributes
    ----------
    credentials_file (str):  
        Path to the JSON credentials file.  
    token_file (str):  
        Path to the token file for caching access credentials.  
    service (googleapiclient.discovery.Resource):  
        Authorized Google Drive API service object, created during initialization 
        and reused for all Drive requests.  

    Methods
    -------
    _authenticate():  
        Authenticate with Google Drive using credentials and return a service object. 
        Handles refreshing and caching tokens automatically.  
    download_file(file_id (str), output_path (str)) -> (bool, str | None):  
        Download a file from Google Drive by its file ID and save it locally at the given path.  
    upload_file(file_id (str), file_path (str)) -> (bool, str | None):  
        Upload (update) a local file to an existing file on Google Drive, specified by its file ID.  
    """

    # Full access to user's Google Drive
    SCOPES = ['https://www.googleapis.com/auth/drive']

    def __init__(self, credentials_file='credentials.json', token_file='token.json'):
        """ Initialize the Google Drive client helper. See class docstring for parameter/attribute details."""

        self.credentials_file = credentials_file
        self.token_file = token_file
        self.service = self._authenticate()

    def _authenticate(self):
        """
        Handles authentication and returns an authorized Google Drive API service object to access the givne Drive.

        - Loads cached credentials from token_file if available.
        - Refreshes the token if expired.
        - Otherwise runds 0Auth with credentials_file and saves the token for future sessions.
        """

        creds = None

        # Load cached credentials if they exist
        if os.path.exists(self.token_file):
            with open(self.token_file, 'rb') as token:
                creds = pickle.load(token)

        # If no valid credentials, either refresh them or start 0Auth flow
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                # Refresh the expired token
                creds.refresh(Request())
            else:
                # Start a new 0Auth flow. Opens the browser for user to login.
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, self.SCOPES)
                creds = flow.run_local_server(port=0)

            # Save the token locally for reuse next session
            with open(self.token_file, 'wb') as token:
                pickle.dump(creds, token)

        # Build and return the Google Drive API client
        return build('drive', 'v3', credentials=creds)

    def download_file(self, file_id, output_path):
        """
        Download a file from Google Drive using its file ID and save it locally.

        :param file_id (str): The Google Drive file ID to download.
        :param output_path (str): Local path where the file is saved.
        :return: (True, None) on success, or (False, error_message) on fail.
        """

        # Return error if no file_id is present
        if not file_id:
            return False, "MISSING_ID"
        
        try:
            # Create a request to download the file
            request = self.service.files().get_media(fileId=file_id)

            print(f"⏳ Downloading {output_path} from Google Drive... please wait.")

            # Use a memory buffer to hold the download
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)

            # Download in chunks until complete
            done = False
            while not done:
                status, done = downloader.next_chunk()
                print(f"Download {int(status.progress() * 100)}%")

            # Write the buffer contents to a local file
            with open(output_path, 'wb') as f:
                f.write(fh.getbuffer())

            print(f"✅ Download complete: {output_path}")
            return True, None
        
        # Handle Google Drive API errors by status code
        except HttpError as e:
            if e.resp.status == 404:
                return False, "NOT_FOUND"  # File does not exist
            elif e.resp.status == 403:
                return False, "PERMISSION_DENIED"  # User does not have permission
            else:
                return False, f"HTTP_ERROR: {e}"  # Other errors

    def upload_file(self, file_id, file_path):
        """
        Upload (update) a local file to an existing file on Google Drive.

        :param file_id: The Google Drive file ID of the file to update.
        :param file_path: Local path of the file to upload.
        :return: (True, None) on success, or (False, error_message) on failure.
        """

        # Return error if no file_id is present
        if not file_id:
            return False, "MISSING_ID"

        try:
            # Create a media upload object for the file
            media = MediaFileUpload(
                file_path,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                resumable=True
            )

            print(f"⏳ Uploading {file_path} to Google Drive... please wait.")

            # Update the existing Drive file with the new content
            self.service.files().update(fileId=file_id, media_body=media).execute()

            print(f"✅ File updated on Google Drive (ID: {file_id})")
            return True, None
        
        # Handle Google Drive API errors by status code
        except HttpError as e:
            if e.resp.status == 404:
                return False, "NOT_FOUND"  # File does not exist
            elif e.resp.status == 403:
                return False, "PERMISSION_DENIED"  # User does not have permission
            else:
                return False, f"HTTP_ERROR: {e}"  # Other errors
            
class ExcelHelper:
    """
    Manage Excel file associated tasks including opening, creating a DataFrame, Updating and restoring column width.

    Methods
    -------
    open_excel_file(filepath (str)) -> None:  
        Open an Excel file with the system’s default application.  
    create_df(file_path (str), sheet_name (str, optional): Default = 'Sheet1') -> (pd.DataFrame, dict):  
        Load an Excel sheet into a DataFrame. 
    update_single_sheet(file_path (str), df (pd.DataFrame), sheet_name (str)) -> None:  
        Overwrite a single sheet in an Excel file with new DataFrame contents, preserving other sheets.  
    """
    
    @staticmethod
    def open_excel_file(filepath):
        """
        Open an Excel file from the given filepath.

        :param file_path (str): Path to the Excel file.
        """

        system = platform.system()
        if system == "Windows":
            # Open with default app in windows
            os.startfile(filepath)
        elif system == "Darwin":
            # USe 'open' command for macOS
            subprocess.call(["open", filepath])
        elif system == "Linux":
            # USe 'xdg-open' command for linux
            subprocess.call(["xdg-open", filepath])
        else:
            # Error message if OS is not supported
            print(f"Unsupported OS: {system}")
    
    @staticmethod
    def create_df(file_path, sheet_name='Sheet1'):
        """
        Load an Excel sheet into a pandas DataFrame.

        - Converts datetime columns into "mm/dd/yyyy" formatted strings.

        :param file_path (str): Path to the Excel file.
        :param sheet_name (str): Name of the sheet to load. Defaults to 'Sheet1'.
        :return: (DataFrame, dict) -> DataFrame of sheet contents.
        """

        # Load sheet into pandas DataFrame
        df = pd.read_excel(file_path, sheet_name)

        return df

    @staticmethod
    def update_single_sheet(file_path, df, sheet_name):
        """
        Overwrite a single sheet in an Excel file with new DataFrame contents.

        - Creates the workbook if it does not exist.
        - Replaces all rows in the specified sheet with new DataFrame data.
        - Preserves all other sheets

        :param file_path (str): Path to the Excel file.
        :param df (pd.DataFrame): The DataFrame to write into the sheet.
        :param sheet_name (str): Name of the sheet to update. Defaults to 'Sheet1'.
        """

        # Create a new workbook if the file doesn't exist and name the sheet sheet_name
        if not os.path.exists(file_path):
            wb = Workbook()
            wb.remove(wb.active)  # Remove default sheet
            wb.create_sheet(title=sheet_name)
            wb.save(file_path)

        wb = load_workbook(file_path)

        # Ensure that the sheet exists
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet {sheet_name} not found in workbook")

        ws = wb[sheet_name]

        # Delete all rows to clear sheet
        ws.delete_rows(1, ws.max_row)

        # Append DataFrame contents row by row including the header
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        wb.save(file_path)
        wb.close()

class IDManager:
    """
    A helper class to manage Google Drive file IDs used in the application.

    - Loads IDs from a JSON config file (or uses defaults if missing).  
    - Allows updating, resetting, and retrieving IDs.  
    - Provides a Tkinter popup for user-friendly ID correction.  

    Parameters
    ----------
    config_file (str, optional): Default = "drive_ids.json"  
        Path to the JSON file where IDs are stored.  
    default_ids (dict, optional): Default = None  
        Dictionary of default keys/IDs used if the config file does not exist.  

    Attributes
    ----------
    config_file (str):  
        Path to the JSON file where IDs are stored.  
    default_ids (dict):  
        Dictionary of default keys/IDs.  
    _id_registry (dict):  
        Stores the file IDs in memory, loaded from the config file or defaults.  

    Methods
    -------
    get_id(key (str)) -> str | None:  
        Retrieve a stored ID by its key (e.g., "freezer80").  
    update_id(key (str), new_id (str)) -> None:  
        Update a single ID in memory and persist the change to the JSON file.  
    get_all_ids() -> dict:  
        Return a shallow copy of all stored IDs.  
    change_id_window(parent (tk.Widget), key (str)) -> None:  
        Open a Tkinter popup window to update a file ID interactively.  
    """

    def __init__(self, config_file="drive_ids.json", default_ids=None):
        """ Initializes the IDManager. See class dostring for parameter/attribute details"""
        
        if default_ids is None:
            default_ids = {
                "grid_dewar": "",
                "freezer80": "",
                "freezer20": "",
                "cell_culture_rows": "",
                "cell_culture_grid": ""
            }
        self.config_file = config_file
        self.default_ids = default_ids
        # Load IDs from file or initialize with defaults
        self._id_registry = self._load_ids()

    def _load_ids(self):
        """Load IDs from JSON file, or use defaults if file doesn't exist."""
        
        # Get file IDs from config_file and return them
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                return json.load(f)
        # If no file, save defaults and return them
        else:
            self._save_ids(self.default_ids)
            return self.default_ids.copy()

    def _save_ids(self, ids):
        """Save IDs dictionary to JSON file."""

        with open(self.config_file, "w") as f:
            json.dump(ids, f, indent=4)

    def get_id(self, key):
        """Retrieve an ID by its key (Ex: 'freezer80')."""

        return self._id_registry.get(key)

    def update_id(self, key, new_id):
        """
        Update a single ID in memory and persist to file.
        
        :param key (str): The ID key to update (Ex: 'freezer80').
        :param new_id (str): The new Google Drive file ID.
        """
        self._id_registry[key] = new_id
        self._save_ids(self._id_registry)

    def get_all_ids(self):
        """Return a copy of all stored IDs (dict)."""

        return self._id_registry.copy()
    
    def change_id_window(self, parent, key):
        """
        Display a Tkinter popup to let the user update a missing/invalid Google Drive ID.

        :param parent: The parent Tkinter window.
        :param key (str): The ID key being updated (e.g., 'freezer80').
        """

        # Create a new popup window using ToplevelWindowHelper
        ID_change_helper = ToplevelWindowHelper(parent, "Update Google Drive ID", size="700x500")
        frame = ID_change_helper.get_main_frame()

        # Instruction label
        tk.Label(
            frame,
            text="The Google Drive ID is missing or invalid."
                "Find the file on Google Drive, copy the ID from the URL (between '/d/' and next '/'),"
                "and enter it below:",
            font=("Arial", 14),
            wraplength=500,
            justify="center"
        ).grid(row=0, column=0, padx=40, pady=(20, 30))

        # Show current incorrect ID and its Key
        tk.Label(
            frame,
            text=f"Incorrect ID: {key}\nCurrent ID: {self.get_id(key)}",
            font=("Arial", 12),
            justify="center"
        ).grid(row=1, column=0, padx=20, pady=(10, 20))

        # Entry field (widget) for new ID
        tk.Label(frame, text="Enter New Google Drive ID:", font=("Arial", 12)).grid(row=2, column=0, sticky="w", padx=20)
        id_entry = tk.Entry(frame, font=("Arial", 12), width=50)
        id_entry.grid(row=3, column=0, padx=20, pady=(5, 20))

        # Restart the program after updating
        def restart_program():
            app_context.temp_file_manager.cleanup_temp_files()
            python = sys.executable
            os.execv(python, [python] + sys.argv)

        # Handle update button click
        def on_update():
            new_id = id_entry.get().strip()
            if new_id:
                # Save new ID
                self.update_id(key, new_id)
                print(f"✅ Updated ID for '{key}' to {new_id}")
                messagebox.showinfo("Success", f"Google Drive ID updated for '{key}'.\nPress 'OK' to restart the program.")
                restart_program()
            else:
                messagebox.showerror("Error", "Please enter a valid Google Drive ID.")

        # Update button
        tk.Button(frame, text="Update & Retry", command=on_update, font=("Arial", 12)).grid(
            row=4, column=0, padx=100, pady=(10, 10), sticky="ew"
        )
        frame.grid_columnconfigure(0, weight=1)