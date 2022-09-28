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

        # Connect to the Gmail API
    serviceGmail = build("gmail", "v1", credentials=creds)

    # We get only messages with the label ID "Label_5005393637772983410"
    # (the "Airbnb Automation" label)
    result = (
        serviceGmail.users()
        .messages()
        .list(labelIds=["Label_5005393637772983410"], userId="me")
        .execute()
    )

    # Previously, we were checking for UNREAD for the messages labeled "Airbnb Automation"
    # but no longer need to because we are moving each message to the Trash in each iteration
    # result =
    # serviceGmail.users().messages().list(labelIds=['Label_5005393637772983410','UNREAD'], userId='me').execute()
    messages = result.get("messages")

    try:
        if messages:  # If there are any messages, iterate else sys.exit()
            # iterate through all the messages
            for msg in messages:  # main iteration
                # Get the message from its id
                txt = (
                    serviceGmail.users()
                    .messages()
                    .get(userId="me", id=msg["id"])
                    .execute()
                )

                # Use try-except to avoid any Errors
                try:
                    # Get value of 'payload' from dictionary 'txt'
                    payload = txt["payload"]
                    headers = payload["headers"]  # Subject is in the headers

                    # Look for Subject in the headers
                    for d in headers:
                        if (
                            d["name"] == "Subject"
                        ):  # this iterates thru the headers until it finds the subject header
                            subject = d[
                                "value"
                            ]  # then puts Subject's value in subject variable
                            break

                    if (
                        "Update Income Report#" in subject
                    ):  # only keep going if it's an Income Report otherwise go to beginning of loop to next message
                        pass
                    else:
                        continue

                    # Remove what we don't need from the Subject
                    subject = subject.replace("Update Income Report#", "")

                    # Put the reservation info we need in multiple variables
                    (
                        cal_checkin,
                        cal_nights,
                        cal_name,
                        cal_total,
                        cal_airbnb_listing_id,
                        cal_cleaning_fee,
                    ) = subject.split("|")

                    cal_nights = cal_nights.replace(" nights", "")

                    cal_total = get_total_after_cleaning_and_tax(
                        cal_total,
                        cal_airbnb_listing_id,
                        cal_cleaning_fee,
                    )
                    cal_checkin = format_reservation_iso(cal_checkin)
                    gsheets_id = get_sheets_id(cal_airbnb_listing_id)

                    # Getting Sheet (tab) name and range info that we need to use
                    gsheets_tab = f"{datetime.strptime(cal_checkin, '%Y-%m-%d').date().strftime('%B %Y')}"
                    gsheets_range = f"{gsheets_tab}!A1:D15"

                    values = [
                        [cal_checkin, cal_nights, cal_name, cal_total],
                    ]
                    body = {"values": values}

                    # Connect to the Google Sheets API
                    serviceSheets = build("sheets", "v4", credentials=creds)

                    # Finding Row number of Guest Name in Column C so we can update the info
                    response = (
                        serviceSheets.spreadsheets()
                        .values()
                        .get(spreadsheetId=gsheets_id, range=f"{gsheets_tab}!C1:C15")
                        .execute()
                    )

                    # Looking for the row number within the name/reservation
                    for _value in response["values"]:
                        if _value[0] == cal_name:
                            row_with_name = response["values"].index(_value) + 1
                            break

                    gsheets_range = f"{gsheets_tab}!A{row_with_name}:D{row_with_name}"

                    # Insert the info
                    serviceSheets.spreadsheets().values().update(
                        spreadsheetId=gsheets_id,
                        range=gsheets_range,
                        valueInputOption="USER_ENTERED",
                        body=body,
                    ).execute()

                    # Get the current sheet's ID (gid)
                    spreadsheet = (
                        serviceSheets.spreadsheets()
                        .get(spreadsheetId=gsheets_id)
                        .execute()
                    )
                    gsheet_id = None
                    for _sheet in spreadsheet["sheets"]:
                        if _sheet["properties"]["title"] == gsheets_tab:
                            gsheet_id = _sheet["properties"]["sheetId"]
                            break

                    requests = [
                        {
                            "sortRange": {
                                "range": {
                                    "sheetId": gsheet_id,
                                    "startRowIndex": 1,
                                    "endRowIndex": 15,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": 4,
                                },
                                "sortSpecs": [{"sortOrder": "ASCENDING"}],
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": gsheet_id,
                                    "dimension": "COLUMNS",
                                    "startIndex": 0,
                                    "endIndex": 1,
                                },
                                "properties": {"pixelSize": 120},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": gsheet_id,
                                    "dimension": "COLUMNS",
                                    "startIndex": 1,
                                    "endIndex": 2,
                                },
                                "properties": {"pixelSize": 85},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": gsheet_id,
                                    "dimension": "COLUMNS",
                                    "startIndex": 2,
                                    "endIndex": 3,
                                },
                                "properties": {"pixelSize": 265},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": gsheet_id,
                                    "dimension": "COLUMNS",
                                    "startIndex": 3,
                                    "endIndex": 4,
                                },
                                "properties": {"pixelSize": 200},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": gsheet_id,
                                    "dimension": "COLUMNS",
                                    "startIndex": 4,
                                    "endIndex": 5,
                                },
                                "properties": {"pixelSize": 110},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": gsheet_id,
                                    "dimension": "COLUMNS",
                                    "startIndex": 5,
                                    "endIndex": 6,
                                },
                                "properties": {"pixelSize": 110},
                                "fields": "pixelSize",
                            }
                        }, #below is for auto-resize but doesn't seem to work to well
                        # {
                        #     "autoResizeDimensions": {
                        #         "dimensions": {
                        #             "sheetId": gsheet_id,
                        #             "dimension": "COLUMNS",
                        #             "startIndex": 0,
                        #             "endIndex": 6,
                        #         }
                        #     }
                        # }
                    ]

                    sort_body = {"requests": requests}

                    serviceSheets.spreadsheets().batchUpdate(
                        spreadsheetId=gsheets_id, body=sort_body
                    ).execute()

                    # Move the message to the Trash after we are done with it.
                    serviceGmail.users().messages().trash(
                        userId="me", id=msg["id"]
                    ).execute()
                    # Or below will permanently delete the message when we are done with it.
                    # serviceGmail.users().messages().delete(userId='me', id=msg['id']).execute()

                except:
                    pass
                # except HttpError as err:
                # print(err)
    except SystemExit:
        pass


def format_reservation_iso(reservation_date):
    # Find year of reservation by first, splitting up today's date and
    # making the month an int to later compare with reservation months
    date_today = datetime.now().date()
    year_today, month_today, day_today = str(date_today).split("-")
    month_today, year_today, day_today = (
        int(month_today),
        int(year_today),
        int(day_today),
    )

    # Hospitable.com now added the year to the check-in/out dates if the reservation is for the following year
    # which is a good but we need to remove it here because our script doesnt expect the year.
    # Example: Tuesday, January 10 2023
    # Check if there is a year in the date and removes the year if so

    # Python 3.8+ with Walrus operator version below
    # if reservation_date := re.sub(r" +[2][0][0-9]{2}", "", reservation_date):
    # pass

    # Pre Python 3.8 version below. (needed for current Python 3.7 environment)
    matches = re.search(r"( +[2][0][0-9]{2})", reservation_date)
    if matches:
        reservation_date = reservation_date.replace(matches.group(1), "")

    # split up reservation date so we can compare months to see if date has already passed or not
    # to datetime object
    reservation_date = datetime.strptime(reservation_date, "%A, %B %d").date()
    # now unpack as a string to month and day using strftime
    reservation_month, reservation_day = reservation_date.strftime("%m %d").split(" ")
    # convert month and day to int for comparisons below
    reservation_month, reservation_day = int(reservation_month), int(reservation_day)

    # if the check-in month is behind current month, the check-in year is next year,
    # else it's the current year
    if reservation_month < month_today:
        reservation_year = year_today + 1
    else:
        reservation_year = year_today

    # Format reservation months and days with double-digits for ISO if needed
    if reservation_month < 10:
        reservation_month = format(reservation_month, "02")
    if reservation_day < 10:
        reservation_day = format(reservation_day, "02")

    reservation_date = f"{reservation_year}-{reservation_month}-{reservation_day}"

    return reservation_date


def get_sheets_id(s):  # this gets the correct Google Sheets ID
    if s == "23683545":
        gsheets_id = "1ibYqy6eGuqrhxrR70Qti1DZrS4LdAqCs8z5LER8hy-E"
    elif s == "44290026":
        gsheets_id = "1CUVuN4vupVjsSK3trIJtzJDod9NPPYOlZxysUVCwxcU"
    elif s == "670197674052387267":
        gsheets_id = "1Psa8plAoMUcPxvdv7xENXakNMIBFVa1HUMHERX_nkOA"
    else:
        sys.exit("Problem with Sheets ID function")

    return gsheets_id


def get_total_after_cleaning_and_tax(
    net_total,
    property_id,
    cleaning_fee,
):
    # If total contains a "," remove it along with the "$".
    if "," or "$" in net_total:
        net_total = float(net_total.replace(",", "").replace("$", ""))

    cleaning_fee = float(cleaning_fee.replace("$", ""))

    # Deduct the tax withholding from the total on certain properties first
    # It's never exact but at the moment 9-23-2022
    # MX withholding VAT is ~7.1% MX withholding income is ~3.55%
    # combining them and using 10.665% is accurate
    if property_id == "44290026" or property_id == "670197674052387267":
        net_total -= round((net_total * 0.10665), 2)

    net_total -= cleaning_fee

    return f"{net_total:.02f}"


if __name__ == "__main__":
    main()
