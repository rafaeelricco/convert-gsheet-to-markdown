"""
            _                      
   _____   (_)  _____  _____  ____ 
  / ___/  / /  / ___/ / ___/ / __ \
 / /     / /  / /__  / /__  / /_/ /
/_/     /_/   \___/  \___/  \____/ 
                                   
© r1cco.com
                            
Google Sheets to Markdown Converter Module

This module provides functionality to convert Google Sheets data into formatted Markdown tables.
It handles complex spreadsheet elements and uses AI assistance for optimal formatting.

Key Features:
1. Google Sheets API integration with secure authentication
2. Intelligent data formatting using Gemini AI
3. Special elements handling (formulas, dropdowns, checkboxes)
4. Real-time progress tracking with visual feedback
5. AI-assisted file naming
6. Markdown table generation with proper alignment and formatting

The module preserves the original sheet layout while creating clean, readable markdown output
suitable for documentation and content management systems.
"""

import json
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv
import pickle
import google.generativeai as genai
import sys
import time
import random
import threading

os.environ["GRPC_PYTHON_LOG_LEVEL"] = "error"
load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


class ProgressBar:
    """
    A console-based progress bar utility for operation tracking.

    This class provides both deterministic and simulated progress tracking with
    support for multi-step operations and background progress simulation.

    Attributes:
        total_width (int): Visual width of the progress bar in characters
        current_step (str): Description of the current operation
        progress (float): Current completion percentage
        _stop_fake_progress (bool): Control flag for simulated progress
        _is_running_fake (bool): Status flag for simulated progress
        _current_thread (Thread): Background thread for progress simulation
        _step_printed (bool): Track if step description has been displayed
    """

    def __init__(self, total_width=80):
        """
        Initialize the progress bar with specified width.

        Args:
            total_width (int): Width of the progress bar in console characters
        """
        self.total_width = total_width
        self.current_step = ""
        self.progress = 0
        self._stop_fake_progress = False
        self._is_running_fake = False
        self._current_thread = None
        self._step_printed = False

    def update(self, step, progress=None):
        """
        Update the progress bar state and display.

        Args:
            step (str): Description of the current operation step
            progress (float, optional): Completion percentage (0-100)
        """
        if step != self.current_step:
            if self.current_step and self.progress < 100:
                self.update(self.current_step, 100)
            sys.stdout.write(f"\n==> {step}\n")
            self.current_step = step
            self.progress = 0
            self._step_printed = True

        if progress is not None:
            self.progress = min(100, max(0, progress))
            filled_width = int(self.total_width * self.progress / 100)
            empty_width = self.total_width - filled_width
            bar = "#" * filled_width + "-" * empty_width
            sys.stdout.write(f"\r{bar} {self.progress:.1f}%")
            sys.stdout.flush()

            if self.progress == 100:
                sys.stdout.write("\n")
                sys.stdout.flush()
                self._stop_fake_progress = True

    def simulate_progress(self, step, start_from=0, until=80):
        """
        Start simulated progress tracking in the background.

        Args:
            step (str): Description of the operation step
            start_from (float): Initial progress percentage
            until (float): Target progress percentage
        """
        if self._current_thread and self._current_thread.is_alive():
            self._stop_fake_progress = True
            self._current_thread.join()

        self._stop_fake_progress = False
        self._is_running_fake = True
        self.update(step, start_from)

        def fake_progress_worker():
            current_progress = start_from
            while not self._stop_fake_progress and current_progress < until:
                time.sleep(random.uniform(0.5, 2))
                if not self._stop_fake_progress:
                    current_progress += random.uniform(0.5, 2)
                    self.update(step, current_progress)
            self._is_running_fake = False

        self._current_thread = threading.Thread(target=fake_progress_worker)
        self._current_thread.daemon = True
        self._current_thread.start()

    def wait_for_fake_progress(self):
        """Wait for any ongoing simulated progress to complete."""
        if self._current_thread and self._current_thread.is_alive():
            self._current_thread.join()


def get_sheet_metadata(progress):
    """
    Retrieve spreadsheet metadata from JSON configuration.

    Args:
        progress (ProgressBar): Progress tracking instance

    Returns:
        dict: Spreadsheet metadata including ID and title

    Raises:
        FileNotFoundError: If the metadata JSON file is missing
        json.JSONDecodeError: If the metadata file is invalid
    """
    progress.simulate_progress("Loading spreadsheet metadata...")
    json_path = os.path.join("json", "sheet_info.json")
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    progress.update("Loading spreadsheet metadata", 100)
    progress.wait_for_fake_progress()
    return data


def get_sheet_data(service, spreadsheet_id, sheet_title, progress):
    """
    Fetch and format all data from the specified Google Sheet.

    Args:
        service: Google Sheets API service instance
        spreadsheet_id (str): Target spreadsheet identifier
        sheet_title (str): Name of the worksheet to process
        progress (ProgressBar): Progress tracking instance

    Returns:
        list: Formatted sheet data including formulas and validation rules

    Raises:
        Exception: If sheet access or data retrieval fails
    """
    try:
        progress.simulate_progress("Retrieving spreadsheet data...")
        result = (
            service.spreadsheets()
            .get(
                spreadsheetId=spreadsheet_id, ranges=[sheet_title], includeGridData=True
            )
            .execute()
        )
        progress.update("Retrieving spreadsheet data...", 85)

        sheet_data = result["sheets"][0]["data"][0]
        rows = sheet_data.get("rowData", [])
        formatted_data = []

        for row in rows:
            if "values" not in row:
                continue

            row_data = []
            for cell in row["values"]:
                display_value = cell.get("formattedValue", "")
                formula = cell.get("userEnteredValue", {}).get("formulaValue", "")
                data_validation = cell.get("dataValidation", {})
                dropdown_options = data_validation.get("condition", {}).get(
                    "values", []
                )

                cell_value = display_value
                if formula:
                    cell_value += f" [formula: {formula}]"
                if dropdown_options:
                    options = [
                        opt.get("userEnteredValue", "") for opt in dropdown_options
                    ]
                    cell_value += f" [options: {', '.join(options)}]"

                row_data.append(cell_value)

            formatted_data.append(row_data)

        progress.update("Retrieving spreadsheet data", 100)
        progress.wait_for_fake_progress()
        return formatted_data

    except Exception as e:
        print(f"Error accessing spreadsheet: {e}")
        return None


def format_with_gemini(data, progress):
    """
    Use Gemini AI to format data into a readable and organized markdown structure.

    Args:
        data (list): The spreadsheet data to be formatted
        progress (ProgressBar): Progress tracking instance

    Returns:
        str: Formatted markdown text
        None: If an error occurs during formatting

    Raises:
        Exception: If Gemini API encounters an error
    """
    try:
        progress.simulate_progress(
            "Formatting data with Gemini...", start_from=10, until=90
        )

        genai.configure(api_key=os.getenv("GEMINI_API_KEY"), transport="grpc")
        model = genai.GenerativeModel(
            "gemini-2.0-flash",
            generation_config={
                "max_output_tokens": 2000000,
                "temperature": 0.3,
                "top_p": 0.9,
                "top_k": 40,
            },
            system_instruction="""
    You are a spreadsheet to markdown table conversion specialist. Your main tasks are:

    1. Create markdown tables that preserve the original spreadsheet layout:
       - Use | for column delimitation
       - Use appropriate alignment (:--, :--:, --:)
       - Maintain column proportions when possible
       - Handle merged cells by using empty cells where needed

    2. Handle different data formats:
       a) Formulas and calculations:
          - Extract only the numeric result, removing any formulas
          - For cells containing "[formula: ...]", keep only the number before it
          - Right-align numeric results
       
       b) Status indicators:
          - Use [ ] for empty checkboxes
          - Use [x] for checked checkboxes
          - Convert status words (OK, NOK, POK, NR) maintaining original format
          - Center-align status indicators

       c) Text content:
          - Left-align regular text
          - Preserve special characters and formatting
          - Handle multi-line content appropriately

       d) Numerical data:
          - Right-align numbers
          - Maintain decimal places and formatting
          - Preserve percentage values

    3. Table structure:
       - Keep original headers and subheaders
       - Preserve data hierarchy and relationships
       - Handle side-by-side tables appropriately
       - Use empty cells (|  |) for spacing and layout

    4. Special considerations:
       - Preserve summary sections and totals
       - Maintain any notes or annotations
       - Keep original data grouping
       - Handle non-standard layouts when present

    Example of formula handling:
    Input:  "13 [formula: =COUNTIF('ADM-Administradores'!H:H;"OK")]"
    Output: "13"

    Return only the formatted markdown tables, without additional text, explanations or formulas.
    Keep the visual structure as close as possible to the original spreadsheet.
    """,
        )

        data_str = "\n".join([", ".join(row) for row in data])

        prompt = f"""Here is the spreadsheet data:
        {data_str}

        Please format this data in a more readable and organized way, highlighting important information and removing redundancies. Return the result in Markdown format.
        """

        response = model.generate_content(prompt)
        formatted_data = response.text

        progress.update("Formatting data with Gemini", 100)
        progress.wait_for_fake_progress()
        return formatted_data

    except Exception as e:
        print(f"Error formatting data with Gemini: {e}")
        return None


def authenticate_google(progress):
    """
    Authenticate with Google Sheets API using OAuth2.

    Args:
        progress (ProgressBar): Progress tracking instance

    Returns:
        Credentials: Valid Google OAuth2 credentials

    Raises:
        Exception: If authentication fails or credentials cannot be obtained
    """
    try:
        progress.update("Authenticating with Google", 0)
        creds = None
        json_dir = "json"
        
        # Create json directory if it doesn't exist
        if not os.path.exists(json_dir):
            os.makedirs(json_dir)
        
        token_path = os.path.join(json_dir, "token.pickle")

        if os.path.exists(token_path):
            try:
                with open(token_path, "rb") as token:
                    creds = pickle.load(token)
                progress.update("Authenticating with Google", 30)
            except Exception as e:
                print(f"Error reading token.pickle: {e}")
                os.remove(token_path)
                creds = None

        if not creds or not creds.valid:
            progress.update("Authenticating with Google", 50)
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    print(f"Error refreshing token: {e}")
                    creds = None

            if not creds:
                # Use environment variables directly
                client_config = {
                    "installed": {
                        "client_id": os.getenv("GOOGLE_CLIENT_ID"),
                        "project_id": os.getenv("GOOGLE_PROJECT_ID"),
                        "auth_uri": os.getenv("GOOGLE_AUTH_URI"),
                        "token_uri": os.getenv("GOOGLE_TOKEN_URI"),
                        "auth_provider_x509_cert_url": os.getenv("GOOGLE_AUTH_PROVIDER_CERT_URL"),
                        "client_secret": os.getenv("GOOGLE_CLIENT_SECRET"),
                        "redirect_uris": ["http://localhost:8080"]
                    }
                }

                flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                print("\nStarting authentication process...")
                print("Add the following redirect URI in Google Cloud Console:")
                print("http://localhost:8080/")
                creds = flow.run_local_server(port=8080)

                with open(token_path, "wb") as token:
                    pickle.dump(creds, token)

            progress.update("Authenticating with Google", 90)

        progress.update("Authenticating with Google", 100)
        return creds

    except Exception as e:
        progress.update("Authentication error", 100)
        raise Exception(f"Google authentication error: {str(e)}")


def generate_file_name_with_ai(data, progress):
    """
    Generate a markdown filename using AI based on spreadsheet content.

    Args:
        data (list): Spreadsheet data to base the filename on
        progress (ProgressBar): Progress tracking instance

    Returns:
        str: Generated filename with .md extension
        str: 'output.md' if an error occurs

    Note:
        The generated filename will be lowercase, use underscores,
        and contain only alphanumeric characters.
    """
    try:
        progress.simulate_progress("Generating file name...", start_from=0, until=90)

        genai.configure(api_key=os.getenv("GEMINI_API_KEY"), transport="rest")
        model = genai.GenerativeModel(
            "gemini-1.5-pro",
            generation_config={
                "max_output_tokens": 100,
                "temperature": 0.5,
                "top_p": 0.95,
            },
        )

        sample_data = str(data[:3]) if len(data) > 3 else str(data)

        prompt = f"""
        Generate a simple markdown filename based on this spreadsheet data:
        {sample_data}

        Requirements:
        - Use only lowercase letters, numbers and underscores
        - Must end with .md
        - Maximum 50 characters
        - No special characters
        - No spaces
        """

        response = model.generate_content(prompt)
        file_name = response.text.strip().lower()

        if not file_name.endswith(".md"):
            file_name += ".md"

        file_name = "".join(c for c in file_name if c.isalnum() or c in ["_", "."])

        progress.update("Generating filename", 100)
        progress.wait_for_fake_progress()

        return file_name

    except Exception as e:
        progress.update("Error generating filename", 100)
        print(f"\nError generating filename: {e}")
        return "output.md"


def mark_identifiers(data):
    """
    Mark special spreadsheet elements with Markdown formatting.

    Args:
        data (list): Raw spreadsheet data

    Returns:
        list: Data with special elements marked in Markdown format

    Note:
        Handles formulas, dropdowns, and checkboxes with appropriate
        Markdown syntax.
    """
    formatted_data = []
    for row in data:
        formatted_row = []
        for cell in row:
            if "[formula:" in cell:
                cell = f"`{cell}`"

            if "[options:" in cell:
                parts = cell.split(" [options: ")
                value = parts[0]
                options = parts[1].rstrip("]")
                cell = f"{value} <select>{options}</select>"

            if cell.upper() in ["TRUE", "FALSE", "VERDADEIRO", "FALSO"]:
                is_checked = cell.upper() in ["TRUE", "VERDADEIRO"]
                cell = "- [x]" if is_checked else "- [ ]"

            formatted_row.append(cell)
        formatted_data.append(formatted_row)
    return formatted_data


def main():
    """
    Main function that orchestrates the sheet-to-markdown conversion process.

    This function:
    1. Initializes progress tracking
    2. Authenticates with Google Sheets
    3. Retrieves and processes sheet data
    4. Formats data using AI
    5. Generates appropriate filename
    6. Saves the formatted markdown output

    Raises:
        Exception: If any step in the process fails
    """
    progress = ProgressBar()

    try:
        # Create output directory if it doesn't exist
        if not os.path.exists("output"):
            os.makedirs("output")
            print("Created output directory")

        creds = authenticate_google(progress)
        service = build("sheets", "v4", credentials=creds)

        sheet_metadata = get_sheet_metadata(progress)
        spreadsheet_id = sheet_metadata["spreadsheet_id"]
        sheet_title = sheet_metadata["sheet_title"]

        sheet_data = get_sheet_data(service, spreadsheet_id, sheet_title, progress)

        if not sheet_data:
            print("No data was retrieved from the spreadsheet.")
            return

        marked_data = mark_identifiers(sheet_data)

        gemini_formatted_data = format_with_gemini(marked_data, progress)

        file_name = generate_file_name_with_ai(sheet_data, progress)
        file_name = file_name.strip().replace(" ", "_")

        with open(f"output/{file_name}", "w", encoding="utf-8") as f:
            f.write(gemini_formatted_data)

    except Exception as e:
        print(f"\nError during script execution: {e}")
    finally:
        progress.wait_for_fake_progress()


if __name__ == "__main__":
    main()
