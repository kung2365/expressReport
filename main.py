from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import readdbf
import moveFile as mf
import drive
import copyFile
import time
import socket


ODsheetID = '1hhzueH6Ha3uzwoQxrXsYBHehj72RbU9fye6nRW2OkbU'
usingDBSheetID = '1WY9POkGfrSY8pVDJEbsExMI0qIjrBnzV_JYfI4uVlvE'

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']


# The ID and range of a sample spreadsheet.
def getUsingOD():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """

    SAMPLE_RANGE_NAME = 'Sheet1!A1:B'
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
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=usingDBSheetID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    ODList = []
    print(result.get('body'))

    if not values:
        print('No data found.')
    else:
        # print('Name, Major:')
        for row in values:
            if row:
                ODList.append(row)
                # print(row[0])
            # Print columns A and E, which correspond to indices 0 and 4.
            # print(row)
    return ODList


def getUsingPR():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """

    SAMPLE_RANGE_NAME = 'Sheet1!B1:B'
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
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=usingDBSheetID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    PRList = []

    if not values:
        print('No data found.')
    else:
        # print('Name, Major:')
        for row in values:
            if row:
                PRList.append(row[0])
                # print(row[0])
            # Print columns A and E, which correspond to indices 0 and 4.
            # print(row)
    return PRList


def appendRow(sheetID, data):
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

    # Call the Sheets API

    range_ = 'Sheet1!A1:F'
    value_input_option = ''
    insert_data_option = ''

    value_range_body = {
        data
    }

    request = service.spreadsheets().values().append(spreadsheetId=sheetID, range=range_,
                                                     valueInputOption=value_input_option,
                                                     insertDataOption=insert_data_option, body=value_range_body)
    response = request.execute()

    # TODO: Change code below to process the `response` dict:
    print(response)


def getODList():
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

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ODsheetID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    ODList = []

    if not values:
        print('No data found.')
    else:
        # print('Name, Major:')
        for row in values:
            ODList.append(str(row[0]))
            # Print columns A and E, which correspond to indices 0 and 4.
            # print(row)
    return ODList


def appendODList(ODName, sheetID):
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
    range_name = 'Sheet1!A1:B'
    values = [
        [
            ODName, sheetID
        ],
        # Additional rows ...
    ]
    body = {
        'values': values
    }
    result = service.spreadsheets().values().append(
        spreadsheetId=usingDBSheetID, range=range_name,
        valueInputOption='RAW', body=body).execute()
    print('{0} cells appended.'.format(result \
                                       .get('updates') \
                                       .get('updatedCells')))


def appendItemList(dataInput, sheetIDInput):
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
    range_name = 'แผ่น1!A1:K'
    values = [dataInput]
    # Additional rows ...

    body = {
        'values': values
    }
    result = service.spreadsheets().values().append(
        spreadsheetId=sheetIDInput, range=range_name,
        valueInputOption='USER_ENTERED', body=body).execute()
    print('{0} cells appended.'.format(result \
                                       .get('updates') \
                                       .get('updatedCells')))


def getSheetData(sheetID):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
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

    service = build('sheets', 'v4', credentials=creds)

    rangeName = 'แผ่น1!A2:I'
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheetID,
                                range=rangeName).execute()
    values = result.get('values', [])

    return values


def buildSheetData(data):
    pass


def updateLink(dataToUpdate, idx):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
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

    service = build('sheets', 'v4', credentials=creds)

    rangeName = 'Sheet1!D' + str(idx + 2)
    # Call the Sheets API
    data = '=HYPERLINK("https://docs.google.com/spreadsheets/d/' + dataToUpdate \
           + '", "เปิดเอกสาร")'

    values = [
        [
            data
        ],
        # Additional rows ...
    ]

    body = {
        'values': values
    }
    result = service.spreadsheets().values().update(
        spreadsheetId=ODsheetID, range=rangeName,
        valueInputOption='USER_ENTERED', body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))

    return values


def writeAllSheet(dataToUpdate, sheetID):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
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

    service = build('sheets', 'v4', credentials=creds)

    rangeName = 'แผ่น1!A2:J'
    # Call the Sheets API
    values = dataToUpdate
    data = [
        {
            'range': rangeName,
            'values': values
        },
        # Additional ranges to update ...
    ]
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': data
    }
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=sheetID, body=body).execute()
    print('{0} cells updated.'.format(result.get('totalUpdatedCells')))

    return values


def is_connected():
    try:
        # connect to the host -- tells us if the host is actually
        # reachable
        socket.create_connection(("www.google.com", 80))
        return True
    except OSError:
        pass
    return False


def writeSheet():
    usingPR = getUsingPR()
    usingOD = []
    usingODSheetID = []
    for i in getUsingOD():
        usingOD.append(i[0])
        usingODSheetID.append(i[1])

    print(usingOD)
    print(usingPR)
    for i in getODList():
        # print(i, '--')
        if not (i in usingOD):
            print(i, '----')
            SheetID2 = drive.newSheet(i)
            appendODList(i, SheetID2)

    usingPR = getUsingPR()
    usingOD = []
    usingODSheetID = []
    for i in getUsingOD():
        usingOD.append(i[0])
        usingODSheetID.append(i[1])
    sheetIDIndex = 0
    for i in getODList():
        idx = 0
        for od in usingOD:
            if i in od:
                sheetID = usingODSheetID[idx]
            idx += 1

        updateLink(sheetID, sheetIDIndex)
        print(sheetID, '>>>>>>>>>>>>>>>>>>>>>>')
        # sheetValue = getSheetData(usingODSheetID[sheetIDIndex])
        # print(sheetValue)
        dataBuffer = []
        for j in readdbf.getPRList(i):
            itemIndex = 0

            for k in readdbf.getItemList(j):
                list_values = [str(v) for v in k.values()]
                list_values.append(i)
                # for l in range()
                # if list_values != sheetValue[itemIndex]:
                print(list_values, '--------')
                dataBuffer.append(list_values)
                itemIndex += 1

                emtyRow = []
                for dummy in list_values:
                    emtyRow.append('')
            dataBuffer.append(emtyRow)
        print(sheetID)
        writeAllSheet(dataBuffer, sheetID)

        sheetIDIndex += 1


def main():
    while True:
        if is_connected():
            if copyFile.copy():
                writeSheet()
                print('Process finish')

        else:
            print('Internet not connect trying again in 60 sec')
        time.sleep(5)


if __name__ == '__main__':
    main()
