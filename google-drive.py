from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from datetime import datetime, date
import pickle, os.path, sys, re

# Define the SCOPES. If modifying it, delete the token.pickle file.
SCOPES = [
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def main():

    # Variable creds will store the user access token.
    # If no valid token found, we will create one.
    creds = None

    # A way to see working directory info
    # cwd = os.getcwd()  # Get the current working directory (cwd)
    # files = os.listdir(cwd)  # Get all the files in that directory
    # print("Files in %r: %s" % (cwd, files))

    # Create the FULL correct directory path
    # cwd = os.getcwd() this doesn't get the FULL path where the file/script is but the below does
    cwd = os.path.dirname(os.path.abspath(__file__))
    the_path = f"{cwd}/"

    # The file token.pickle contains the user access token. Check if it exists
    if os.path.exists(f"{the_path}token.pickle"):

        # Read the token from the file and store it in the variable creds
        with open((the_path + "token.pickle"), "rb") as token:
            creds = pickle.load(token)

    # If credentials are not available or are invalid, ask the user to log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                (the_path + "credentials.json"),
                SCOPES,
            )
            creds = flow.run_local_server(port=0)

        # Save the access token in token.pickle file for the next run
        with open((the_path + "token.pickle"), "wb") as token:
            pickle.dump(creds, token)

###################################################

    cleaning_form_gsheets_id = "1d8M5sx0n8YlAP_Tw6Ra0qM7FBavyFm3jFpQapPmOpcs"

    # Connect to the Google Sheets API
    serviceSheets = build("sheets", "v4", credentials=creds)

    response = (
        serviceSheets.spreadsheets()
        .values()
        .get(spreadsheetId=cleaning_form_gsheets_id, range="Cleaning Service Form!A1:J")
        .execute()
    )

    # Looking for the row number within the name/reservation
    for _value in response["values"]:
        if _value[2] == "Terrazas 302" and _value[3] and _value[9] != "Yes":

            # Get working row number
            current_row_number = response["values"].index(_value) + 1

            cleaning_week_ending_date = _value[1]
            property_name = _value[2]
            water_service_amount = _value[3]
            # to datetime object
            cleaning_folder = datetime.strptime(
                cleaning_week_ending_date, "%m/%d/%Y"
            ).date()
            
            # now unpack as a string to month and year using strftime
            expense_gsheets_tab = cleaning_folder.strftime("%B %Y")
            cleaning_folder_month, cleaning_folder_year = expense_gsheets_tab.split(" ")

            matches = re.search(r"^.+\?id=(.+)$", _value[4])
            if matches:
                water_service_pdf_id = matches.group(1)

            values = [
                ["Yes"],
                    ]
            body = {"values": values}

            gsheets_range = f"Cleaning Service Form!J{current_row_number}"

            # Insert the info
            serviceSheets.spreadsheets().values().update(
                spreadsheetId=cleaning_form_gsheets_id,
                range=gsheets_range,
                valueInputOption="USER_ENTERED",
                body=body,
            ).execute()

            # Connect to the Google Drive API
            serviceDrive = build("drive", "v3", credentials=creds)

            expense_receipts_folder_id, expense_sheets_id = get_expense_drive_ids(property_name)
            ##################################################################
            query = f"parents = '{expense_receipts_folder_id}'"
            response = serviceDrive.files().list(q=query).execute()
            files = response.get("files")
            nextPageToken = response.get("nextPageToken")
            
            # Add to file list if there is another page of files
            while nextPageToken:
                response = serviceDrive.files().list(q=query,pageToken=nextPageToken).execute()
                files.extend(response.get("files"))
                nextPageToken = response.get("nextPageToken")
            
            # Searching main folder (year)
            for _file in response["files"]:
                if _file["name"] == cleaning_folder_year:
                    expense_year_folder_id = _file["id"]
                    break
            ##################################################################
            query = f"parents = '{expense_year_folder_id}'"
            response = serviceDrive.files().list(q=query).execute()
            files = response.get("files")
            nextPageToken = response.get("nextPageToken")

            # Add to file list if there is another page of files
            while nextPageToken:
                response = serviceDrive.files().list(q=query,pageToken=nextPageToken).execute()
                files.extend(response.get("files"))
                nextPageToken = response.get("nextPageToken")

            # Searching sub folder (month)
            for _file in response["files"]:
                if cleaning_folder_month in _file["name"]:
                    expense_month_folder_id = _file["id"]
                    break

            serviceDrive.files().update(
                fileId=water_service_pdf_id,
                addParents=expense_month_folder_id,
                removeParents="1sNcSK9af7iOyIP3wRF1T4frhaQuHNpKUsrgkExBjPHdtZVSm0QClNiYvX3OAg8dXomiJXno3", # Cleaning form Water Service uploads folder
            ).execute()

            ###############################################################

            gsheets_range = f"{expense_gsheets_tab}!A1:D20"

            values = [
                ["Water Service", cleaning_week_ending_date, water_service_amount, "Chris"],
            ]
            body = {"values": values}

            # Find last row with data in sheet
            first_empty_row = 0

            response = (
                serviceSheets.spreadsheets()
                .values()
                .get(spreadsheetId=expense_sheets_id, range=gsheets_range)
                .execute()
            )

            first_empty_row += len(response["values"]) + 1
            gsheets_range = (
                f"{expense_gsheets_tab}!A{first_empty_row}:D{first_empty_row}"
            )

            # Insert the info
            serviceSheets.spreadsheets().values().update(
                spreadsheetId=expense_sheets_id,
                range=gsheets_range,
                valueInputOption="USER_ENTERED",
                body=body,
            ).execute()
        else:
            continue
        ###############################################################

def get_expense_drive_ids(s):
    if s == "Terrazas 302":
        expense_receipts_folder_id = "1-1ZNm2zutHfohDODCcaEO6xikExA3asp"
        expense_sheets_id = "1DG27ULzS92jK1ZYeKPZeMlJBlJ9vJP5OuK8MGR3XBHw"
    else:
        sys.exit("Problem with Sheets ID function")

    return expense_receipts_folder_id, expense_sheets_id


if __name__ == "__main__":
    main()
