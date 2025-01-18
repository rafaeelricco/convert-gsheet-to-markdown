"""
            _                      
   _____   (_)  _____  _____  ____ 
  / ___/  / /  / ___/ / ___/ / __ \
 / /     / /  / /__  / /__  / /_/ /
/_/     /_/   \___/  \___/  \____/ 
                                   
© r1cco.com

Google Sheets Access Module

This module provides functionality to authenticate with Google Sheets API and extract information
from a specified spreadsheet. It handles OAuth2 authentication, manages credentials, and allows
users to select and save information about specific worksheets.
"""

import gspread
import json
import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def authenticate_google_sheets():
    """
    Authenticate with Google Sheets using OAuth2 credentials from a JSON file.

    This function handles the OAuth2 authentication flow:
    1. Checks for existing credentials in token.pickle
    2. Refreshes expired credentials if possible
    3. Initiates new authentication flow if needed
    4. Saves credentials for future use

    Returns:
        gspread.Client: Authorized gspread client instance for interacting with Google Sheets.
    """
    creds = None
    token_path = os.path.join("json", "token.pickle")
    client_secrets_path = os.path.join(
        "json",
        "client_secret.json",
    )

    if os.path.exists(token_path):
        with open(token_path, "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            with open(client_secrets_path, "r") as f:
                client_config = json.load(f)

            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            print("Starting authentication process...")
            print("Add the following redirect URI in Google Cloud Console:")
            print("http://localhost:8080/")
            creds = flow.run_local_server(port=8080)

            with open(token_path, "wb") as token:
                pickle.dump(creds, token)

    client = gspread.authorize(creds)
    return client


def list_sheets_and_save_info(spreadsheet_id, output_file):
    """
    Access a spreadsheet, list its worksheets, and save selected worksheet information to a JSON file.

    Args:
        spreadsheet_id (str): The ID of the Google Spreadsheet to access.
        output_file (str): Path to the JSON file where worksheet information will be saved.

    Returns:
        None

    Raises:
        gspread.SpreadsheetNotFound: If the specified spreadsheet ID is invalid or inaccessible.
        ValueError: If user input is invalid during worksheet selection.
    """
    client = authenticate_google_sheets()

    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
    except gspread.SpreadsheetNotFound:
        print(f"Error: Spreadsheet with ID '{spreadsheet_id}' not found.")
        return
    except Exception as e:
        print(f"Error accessing spreadsheet: {e}")
        return

    worksheets = spreadsheet.worksheets()
    print("\nAvailable worksheets:")
    for i, sheet in enumerate(worksheets):
        print(f"{i + 1}. {sheet.title} (ID: {sheet.id})")

    while True:
        try:
            choice = int(input("\nChoose the number of the worksheet to access: "))
            if 1 <= choice <= len(worksheets):
                selected_sheet = worksheets[choice - 1]
                break
            else:
                print("Invalid choice. Please try again.")
        except ValueError:
            print("Invalid input. Please enter a number.")

    sheet_info = {
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": selected_sheet.id,
        "sheet_title": selected_sheet.title,
    }

    with open(output_file, "w") as f:
        json.dump(sheet_info, f, indent=4)

    print(f"\nWorksheet '{selected_sheet.title}' information saved to '{output_file}'.")


if __name__ == "__main__":
    try:
        SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
        if not SPREADSHEET_ID:
            raise ValueError(
                "SPREADSHEET_ID not found in environment variables. Check your .env file."
            )

        OUTPUT_FILE = os.path.join("json", "sheet_info.json")

        list_sheets_and_save_info(SPREADSHEET_ID, OUTPUT_FILE)
    except Exception as e:
        print(f"Error: {str(e)}")
