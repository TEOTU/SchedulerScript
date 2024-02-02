import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = "1yca1wLBj0vRrgbG6x_lg5IQtNqH3GAfE36mJWIefAlU"
RANGES = ['Studying!B2:C1009', 'Studying!E2:I1009']  # Topics names and repetition intervals
CREDS = None


def get_sheet_info():
    global CREDS
    service = build("sheets", "v4", credentials=CREDS)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet
        .values()
        .batchGet(spreadsheetId=SPREADSHEET_ID, ranges=RANGES)
        .execute()
    )
    return result


def get_info():
    values = []
    result = get_sheet_info()
    values_ranges = result.get('valueRanges', [])
    if not values_ranges:
        print("No data found.")
    else:
        for block in range(0, 2):  # values[0] - First columns with topics/subtopics, values[1] - columns with intervals
            values.append(values_ranges[block]['values'])
        return values


def column_to_value():
    column_value = []
    row = 0  # the first row
    for i in range(0, 2):  # go through 2 blocks of values
        column_value.append({})  # for each block create corresponding dictionary
        for data in get_info()[i]:
            column_value[i][row] = data
            row += 1
    return column_value


def connect():
    global CREDS
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        CREDS = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not CREDS or not CREDS.valid:
        if CREDS and CREDS.expired and CREDS.refresh_token:
            CREDS.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            CREDS = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(CREDS.to_json())


def main():
    try:
        connect()
        print(column_to_value())
    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
