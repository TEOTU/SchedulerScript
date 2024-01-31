import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1yca1wLBj0vRrgbG6x_lg5IQtNqH3GAfE36mJWIefAlU"
FIRST_TOPIC_RANGE, FIRST_ROW_RANGE = 2, 1  # All the topics are stored from the row 1(which is second) and value 1(also)
CREDS = None
SHEET_NUM = 2  # The number of the sheet in a spreadsheet
DATA_ENTRY = 0


def cache_read():
    # read last cell checked
    with open('cache.txt', 'r') as file:
        return file.read()


def cache_write(last: str):
    # write the last checked cell
    with open('cache.txt', 'w') as file:
        file.write(last)


def last_checked_row():
    last_row = cache_read()
    if not cache_read():  # if empty
        last_row = 1
        return last_row
    else:
        return last_row


def get_sheet_info():
    global CREDS
    service = build("sheets", "v4", credentials=CREDS)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, includeGridData=True)
        .execute()
    )
    return result


def get_color(row):
    result = get_sheet_info()
    values = result.get('values', [])
    if not values:
        print("No data found.")
    else:
        cell_data = result['sheets'][SHEET_NUM]['data'][DATA_ENTRY]['rowData'][row]['values'][FIRST_TOPIC_RANGE]
        return cell_data.get('effectiveFormat', {}).get("backgroundColor", {})  # return color


def sheet_len():
    result = get_sheet_info()
    values = result.get('values', [])
    if not values:
        print('No data found')
        return
    else:
        return len(values)


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


def set_intervals():
    last_row = last_checked_row()
    for row in range(last_row, sheet_len()):
        color_key = get_color(row).key()
        if color_key == 'green':
            print('green')
        if color_key == 'orange':
            print('orange')
        if color_key == 'red':
            print('red')
        if color_key == 'black':
            print('black')

def main():
    try:
        connect()
        set_intervals()
    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
