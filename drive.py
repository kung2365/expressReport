from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# sheet in R&D\docuents folder
folder_id = '15LwXa784Hkw_ZUjNrqs4KhvGCmZBsCyX'

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.metadata']


def moveTofolder(file_id):
    """Shows basic usage of the Drive v3 API.
        Prints the names and ids of the first 10 files the user has access to.
        """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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

    service = build('drive', 'v3', credentials=creds)

    # file_id = '1GxxBecTsC47qvngYAAI6qfTDNriU2L040Poa-ZCMqIg'
    # folder_id = '15LwXa784Hkw_ZUjNrqs4KhvGCmZBsCyX'
    # Retrieve the existing parents to remove
    file = service.files().get(fileId=file_id,
                               fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    # Move the file to the new folder
    file = service.files().update(fileId=file_id,
                                  addParents=folder_id,
                                  removeParents=previous_parents,
                                  fields='id, parents').execute()
    print('creat sheet success')
    return file_id

# The ID and range of a sample spreadsheet.
def createSheet(sheetName):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """

    SAMPLE_RANGE_NAME = 'Sheet1!A1:A'
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API

    spreadsheet = {
        'properties': {
            'title': str(sheetName)
        }
    }
    spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                                fields='spreadsheetId').execute()
    print('Spreadsheet ID: {0}'.format(spreadsheet.get('spreadsheetId')))

    """Shows basic usage of the Sheets API.
            Prints values from a sample spreadsheet.
            """
    SAMPLE_RANGE_NAME = 'Sheet1!B2:B'
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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

    service = build('sheets', 'v4', credentials=creds)
    range_name = 'แผ่น1!A1:I'
    values = [
        ['ลำดับ', 'เลขที่ PR', 'Part Number','Part Name', 'จำนวน', 'หน่วย', 'ลูกค้า', 'วันที่เปิดเอกสาร']

    ]
    # Additional rows ...

    body = {
        'values': values
    }
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet.get('spreadsheetId'), range=range_name,
        valueInputOption='RAW', body=body).execute()
    print('{0} cells appended.'.format(result \
                                       .get('updates') \
                                       .get('updatedCells')))


    return spreadsheet.get('spreadsheetId')


def newSheet(sheetName):
    return moveTofolder(createSheet(sheetName))



# if __name__ == '__main__':
#     newSheet('testPython3')
