from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
# from datetime import datetime, date
import pickle, os.path, sys

# Define the SCOPES. If modifying it, delete the token.pickle file.
SCOPES = [
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
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
                        "Delete Google Calendar Event#" in subject
                    ):  # only keep going if it's an Event cancellation otherwise go to beginning of loop to next message
                        pass
                    else:
                        continue

                    # Remove what we don't need from the Subject
                    subject = subject.replace("Delete Google Calendar Event#", "")

                    # Put the reservation info we need in multiple variables
                    (
                        cal_name,
                        cal_airbnb_listing_id,
                    ) = subject.split("|")

                    # Go get the correct Google Calendar ID
                    gcal_id = get_calendar_id(cal_airbnb_listing_id)

                    # Connect to the Google Calendar API
                    serviceCalendar = build("calendar", "v3", credentials=creds)

                    # Search Calendar by Name of guest (cal_event_name) for existing Event/Reservation
                    page_token = None
                    while True:
                        events = (
                            serviceCalendar.events()
                            .list(calendarId=gcal_id, q=cal_name, pageToken=page_token)
                            .execute()
                        )
                        if events:
                            for event in events["items"]:
                                # Delete the existing Event
                                serviceCalendar.events().delete(
                                    calendarId=gcal_id, eventId=event["id"]
                                ).execute()

                                # Move the message to the Trash after we are done with it.
                                serviceGmail.users().messages().trash(
                                userId="me", id=msg["id"]
                                ).execute()
                                # Or below will permanently delete the message when we are done with it.
                                # serviceGmail.users().messages().delete(userId='me', id=msg['id']).execute()
                            
                            page_token = events.get("nextPageToken")
                            if not page_token:
                                break
                        else:
                            sys.exit("No matching events to delete")

                except:
                    pass
        else:  # Exit if there are no messages
            sys.exit("No messages")
    except SystemExit:
        pass


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

# print(f"delete test {datetime.now()}")