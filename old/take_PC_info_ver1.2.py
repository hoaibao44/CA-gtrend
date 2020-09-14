import os
import subprocess
import platform
import pickle
import os.path
import requests
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


def G_sheet(in_array):
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # The ID and range of a sample spreadsheet.
    SPREADSHEET_ID = '1Lu0e8YRCHYcTgr0MoT0s-aBUPKefx2brBmr_lb3V-Dc'
    RANGE_NAME = 'All_Staff!A1:G'    

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
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])

    print(values[0])

    values = [in_array]
    body = {
    "majorDimension": "ROWS",
    "values": values
    }
    result = sheet.values().append(spreadsheetId=SPREADSHEET_ID, range='All_Staff!A2:H',
    valueInputOption="USER_ENTERED", body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))

def get_PC_info():
    myArray =[]

    #PC ID
    myArray.append(os.environ['COMPUTERNAME'])

     #PC sevice tag
    myArray.append(subprocess.check_output('wmic csproduct get IdentifyingNumber').decode().split('\n')[1].strip())
    
    #win ver
    myArray.append(platform.platform())

    #user ID
    myArray.append(os.getlogin())

    #hard disk model
    #hard disk serial number

    hard_disk_model=[]
    hard_disk_model = subprocess.check_output('wmic diskdrive get Model').decode().split('\n')
    
    myArray.append(subprocess.check_output('wmic diskdrive get Model').decode().split('\n')[1].strip())
    myArray.append(subprocess.check_output('wmic diskdrive get SerialNumber').decode().split('\n')[1].strip())
    myArray.append(subprocess.check_output('wmic diskdrive get Model').decode().split('\n')[2].strip())
    myArray.append(subprocess.check_output('wmic diskdrive get SerialNumber').decode().split('\n')[2].strip())
    #IP info
    #myArray.append(subprocess.check_output('ipconfig /all').decode('Shift-JIS').strip())

    print(myArray)
    print('get info: DONE')
    return myArray
if __name__ == '__main__':
    pc_info = get_PC_info()
    G_sheet(pc_info)
