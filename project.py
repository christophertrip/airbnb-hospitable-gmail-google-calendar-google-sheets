from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from datetime import datetime, date
import pickle, os.path, sys, re

# Define the SCOPES. If modifying it, delete the token.pickle file.
SCOPES = ["https://mail.google.com/", "https://www.googleapis.com/auth/calendar"]


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

                    # Remove what we don't need from the Subject
                    subject = subject.replace("Google Calendar Script#", "")

                    # Put the reservation info we need in multiple variables
                    (
                        cal_name,
                        cal_total,
                        cal_children,
                        cal_infants,
                        cal_phone,
                        cal_whatsapp,
                        cal_email,
                        cal_checkin,
                        cal_checkout,
                        cal_airbnb_listing_id,
                    ) = subject.split("|")

                    # Check guest quantity and types adults, children, infants to we can get the breakdown correct
                    match1 = re.search(
                        r"([0]\s\bchildren\b|[0]\s\bchild\b)", cal_children
                    )
                    match2 = re.search(
                        r"([0]\s\binfants\b|[0]\s\binfant\b)", cal_infants
                    )
                    if not match1 or not match2 and "1 person" not in cal_total:
                        cal_event_name = f"{cal_name} ({guest_types(cal_total, cal_children, cal_infants)})"
                    elif "person" in cal_total:
                        cal_total = cal_total.replace("person", "adult")
                        cal_event_name = f"{cal_name} ({cal_total})"
                    else:
                        cal_total = cal_total.replace("people", "adults")
                        cal_event_name = f"{cal_name} ({cal_total})"

                    # Put together the final ISO formatted check-in/out dates (YYYY-MM-DD)
                    cal_checkin = format_reservation_iso(cal_checkin)
                    cal_checkout = format_reservation_iso(cal_checkout)

                    # Create Calendar event description from available varaiables
                    cal_description = f"{cal_event_name}\n\n{cal_phone}\n\n{cal_whatsapp}\n\n{cal_email}"

                    # Go get the correct Google Calendar ID
                    gcal_id = get_calendar_id(cal_airbnb_listing_id)

                    # Connect to the Google Calendar API
                    serviceCalendar = build("calendar", "v3", credentials=creds)

                    # Create Calendar Event
                    event_request_body = {
                        "summary": cal_event_name,
                        "description": cal_description,
                        "start": {
                            "dateTime": f"{cal_checkin}T15:00:00-05:00",
                            "timeZone": "America/Cancun",
                        },
                        "end": {
                            "dateTime": f"{cal_checkout}T11:00:00-05:00",
                            "timeZone": "America/Cancun",
                        },
                    }

                    # This creates the Calendar Event
                    serviceCalendar.events().insert(
                        calendarId=gcal_id, body=event_request_body
                    ).execute()

                    # Move the message to the Trash after we are done with it.
                    serviceGmail.users().messages().trash(
                        userId="me", id=msg["id"]
                    ).execute()
                    # Or below will permanently delete the message when we are done with it.
                    # serviceGmail.users().messages().delete(userId='me', id=msg['id']).execute()

                except:
                    pass
        else:  # Exit if there are no messages
            sys.exit("No messages")
    except SystemExit:
        pass


def guest_types(total_people, children, infants):

    total_people = re.sub(r"\s\bpeople\b", "", total_people)
    total_people = int(total_people)

    children = re.sub(r"\s\bchildren\b|\s\bchild\b", "", children)
    children = int(children)

    infants = re.sub(r"\s\binfants\b|\s\binfant\b", "", infants)
    infants = int(infants)

    total_adults = total_people - children - infants

    if total_adults == 1:
        total_adults = f"{total_adults} adult"
    else:
        total_adults = f"{total_adults} adults"

    if children == 0 and infants == 1:
        total_people = f"{total_adults}, {infants} infant"
    elif children == 0 and infants > 1:
        total_people = f"{total_adults}, {infants} infants"
    elif children == 1 and infants == 0:
        total_people = f"{total_adults}, {children} child"
    elif children > 1 and infants == 0:
        total_people = f"{total_adults}, {children} children"
    elif children == 1 and infants == 1:
        total_people = f"{total_adults}, {children} child, {infants} infant"
    elif children == 1 and infants > 1:
        total_people = f"{total_adults}, {children} child, {infants} infants"
    elif children > 1 and infants == 1:
        total_people = f"{total_adults}, {children} children, {infants} infant"
    else:
        total_people = f"{total_adults}, {children} children, {infants} infants"

    return total_people


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
    reservation_date = datetime.strptime(reservation_date, "%A, %B %d").date()
    reservation_month, reservation_day = (
        str(reservation_date).replace("1900-", "").split("-")
    )
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


def get_calendar_id(s):  # this gets the correct Google Calendar ID
    if s == "23683545":
        gcal_id = "lgfmqv4o9q27f4afc3ml3no3ik@group.calendar.google.com"
    elif s == "44290026":
        gcal_id = "fuh78sb1opc86r2jmolflhnmfc@group.calendar.google.com"
    elif s == "670197674052387267":
        gcal_id = "8l321gb0rjua7bt6p89mi34qqc@group.calendar.google.com"
    else:
        gcal_id = "primary"

    return gcal_id


if __name__ == "__main__":
    main()