import os.path
import re
import sys
from datetime import datetime
from string import ascii_uppercase
import io

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_authenticated_service():
    """Authenticate with Google Sheets API and return the service."""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if os.path.exists("credentials.json"):
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            else:
                raise FileNotFoundError(
                    "credentials.json not found, see\
                    https://developers.google.com/workspace/sheets/api/quickstart/python\
                    to set up an OAuth client and download the credentials.json file"
                )
    # Save the credentials for the next run
    with open("token.json", "w") as token:
        token.write(creds.to_json())

    return build("sheets", "v4", credentials=creds)


SAVE_FILE = ".prev.txt"

# Initialize service once at module level
service = get_authenticated_service()


def get_spreadsheet_title(sheet_id):
    """Get the title of a Google Spreadsheet by ID."""
    try:
        result = (
            service.spreadsheets()
            .get(spreadsheetId=sheet_id, fields="properties.title")
            .execute()
        )
        return result.get("properties", {}).get("title")
    except HttpError as err:
        raise RuntimeError(f"Failed to get spreadsheet title: {err}") from err


def get_sheet_names(sheet_id):
    """Get list of all sheet names in a spreadsheet."""
    try:
        result = (
            service.spreadsheets()
            .get(spreadsheetId=sheet_id, fields="sheets.properties.title")
            .execute()
        )
        sheets = result.get("sheets", [])
        return [sheet["properties"]["title"] for sheet in sheets]
    except HttpError as err:
        raise RuntimeError(f"Failed to get sheet names: {err}") from err


def select_sheet(sheet_id):
    """Let user select which sheet to use from the spreadsheet."""
    sheet_names = get_sheet_names(sheet_id)

    if not sheet_names:
        raise ValueError("No sheets found in spreadsheet")

    if len(sheet_names) == 1:
        return sheet_names[0]

    print("\033[97mAvailable sheets:")
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}")

    while True:
        try:
            choice = int(input("Select sheet number:\n→\033[0m ").strip())
            if 1 <= choice <= len(sheet_names):
                return sheet_names[choice - 1]
            print(f"Please enter a number between 1 and {len(sheet_names)}")
        except ValueError:
            print("Invalid input, please enter a number")


def column_letter(n):
    """Convert 0-indexed column number to letter(s) (0->A, 25->Z, 26->AA, etc)."""
    result = ""
    while n >= 0:
        result = ascii_uppercase[n % 26] + result
        n = n // 26 - 1
    return result


def find_date_column(sheet_id, sheet_name):
    """Find column with today's date or return next empty column."""
    today = datetime.now().strftime("%m/%d/%Y")

    try:
        result = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=sheet_id, range=f"{sheet_name}!1:1")
            .execute()
        )
        row = result.get("values", [[]])[0] if result.get("values") else []

        # Check for today's date
        for i, cell in enumerate(row):
            if cell.strip() == today:
                return column_letter(i)

        # Find next empty column
        col_index = len(row)
        return column_letter(col_index)
    except HttpError as err:
        raise RuntimeError(f"Failed to read header row: {err}") from err


def write_date_header(sheet_id, sheet_name, column):
    """Write today's date to the column header (row 1)."""
    today = datetime.now().strftime("%m/%d/%Y")

    try:
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!{column}1",
            valueInputOption="USER_ENTERED",
            body={"values": [[today]]},
        ).execute()
    except HttpError as err:
        raise RuntimeError(f"Failed to write date header: {err}") from err


def get_next_empty_row(sheet_id, sheet_name, column):
    """Find the first empty row in the column (starting from row 2)."""
    try:
        result = (
            service.spreadsheets()
            .values()
            .get(spreadsheetId=sheet_id, range=f"{sheet_name}!{column}2:{column}")
            .execute()
        )
        values = result.get("values", [])
        return len(values) + 2  # +2 because we start counting from row 2
    except HttpError as err:
        raise RuntimeError(f"Failed to find next empty row: {err}") from err


def write_attendance_id(sheet_id, sheet_name, column, row, attendance_id):
    """Write a single attendance ID to a specific cell."""
    try:
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_name}!{column}{row}",
            valueInputOption="USER_ENTERED",
            body={"values": [[attendance_id]]},
        ).execute()
    except HttpError as err:
        raise RuntimeError(f"Failed to write ID to sheet: {err}") from err


def read_and_write_attendance(sheet_id, sheet_name, column):
    """Read 7-digit IDs from stdin and write them one by one to the sheet."""
    print("\033[97mEnter 7-digit attendance IDs (Ctrl+D or Ctrl+Z to finish):\033[0m")

    # Write the date header
    write_date_header(sheet_id, sheet_name, column)

    # Get the next empty row
    row = get_next_empty_row(sheet_id, sheet_name, column)
    count = 0

    try:
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            if not line.isdigit() or len(line) != 7:
                # Move cursor to end of line, print in red, then move to next line
                print(f"\0337\033[1A\033[{len(line)}C\033[31m ✗ Invalid input (must be exactly 7 digits)\033[0m", end="\0338")
            else:
                # Write immediately to sheet
                write_attendance_id(sheet_id, sheet_name, column, row, line)
                print(f"\0337\033[1A\033[{len(line)}C\033[32m ✓ row {row}\033[0m", end="\0338")
                row += 1
                count += 1
            # flush
            sys.stdout.flush()
    except (EOFError, KeyboardInterrupt):
        pass

    print(f"Finished. Wrote {count} IDs.")

if __name__ == "__main__":
    saved = ""
    title = ""
    try:
        with open(SAVE_FILE, "r", encoding="utf-8") as saved_file:
            saved = saved_file.read().strip()
            title = get_spreadsheet_title(saved)
    except FileNotFoundError:
        pass

    sheet_id = ""
    if saved:
        if (
            input(f"\033[97mUse previously saved Google Sheet ({title})? (y/n)\n→ \033[0m")
            .strip()
            .lower()
            == "y"
        ):
            sheet_id = saved

    if not sheet_id:
        sheet_link = input(
            "\033[97mEnter the Google Sheet link (ensure that it is editable by the authorized account)\n→ \033[0m"
        ).strip()

        # extact sheet id
        match = re.search(r"/d/([A-Za-z0-9\-_]{44})", sheet_link)
        if not match:
            raise ValueError("No 44-character id found in the provided link.")
        sheet_id = match.group(1)
        if len(sheet_id) != 44:
            raise ValueError(
                "Extracted id does not have the required length of 44 characters."
            )

    # save the google sheet id
    with open(SAVE_FILE, "w", encoding="utf-8") as saved_file:
        saved_file.write(sheet_id)

    # Let user select which sheet to use
    sheet_name = select_sheet(sheet_id)

    # Find or create column for today
    column = find_date_column(sheet_id, sheet_name)

    # Read and write attendance IDs one by one
    read_and_write_attendance(sheet_id, sheet_name, column)
    
