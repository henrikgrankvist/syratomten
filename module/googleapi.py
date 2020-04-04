# Google API
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

class Google:

    def token():
        """
        The file token.pickle stores the user's access and refresh tokens, and is
        created automatically when the authorization flow completes for the first
        time.
        """
        creds = None

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        return creds

    def create_spreadsheet(spreadsheet_title, sheet_titles):

        creds = Google.token()

        service = build('sheets', 'v4', credentials=creds)

        spreadsheet = {
            "properties": {
                "title": spreadsheet_title
            },
            "sheets": sheet_titles
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet,fields='spreadsheetId').execute()
        #print('Spreadsheet ID: {0}'.format(spreadsheet.get('spreadsheetId')))
        return spreadsheet.get('spreadsheetId')

    def get(spreadsheet_id=None, sheet_name=None, sheet_range='!A2:I'):

        creds = Google.token()

        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=sheet_name+sheet_range).execute()

        return result.get('values', [])

    def update(spreadsheet_id, sheet, range, data):

        creds = Google.token()

        service = build('sheets', 'v4', credentials=creds)

        # data should be an array (list within a list)
        body = {
            "values" : data
        }

        sheet_range = sheet + range

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().update(spreadsheetId=spreadsheet_id, range=sheet_range, valueInputOption='RAW', body=body).execute()

        return result

    def append(spreadsheet_id, sheet, range, data):

        creds = Google.token()

        service = build('sheets', 'v4', credentials=creds)

        body = {
            "values" : data
        }

        sheet_range = sheet + range

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().append(spreadsheetId=spreadsheet_id, range=sheet_range, valueInputOption='RAW', body=body).execute()

        return result
