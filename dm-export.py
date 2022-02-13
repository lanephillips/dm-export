from __future__ import print_function

import os.path
import configparser

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE_NAME = 'Class Data!A2:E'


def loadConfig():
    print('Load settings.')
    config = configparser.ConfigParser()
    config.read('config.ini')
    return config


def saveConfig(config):
    print('Save settings.')
    with open('config.ini', 'w') as configfile:
        config.write(configfile)


def loginGoogle():
    print('Log into Google Sheets.')
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def loginTwitter():
    print('Log into Twitter.')
    pass


def openSpreadsheet(creds):
    print('Open spreadsheet.')
    # TODO: check for spreadsheet ID in the config file
    # TODO: if it's not there prompt the user for a name and create it
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[4]))
    except HttpError as err:
        print(err)


def exportMentions():
    print('Export mentions to spreadsheet.')
    # TODO: some way of avoiding re-export
    pass


def exportDMs():
    print('Export direct messages to spreadsheet.')
    # TODO: some way of avoiding re-export


def main():
    config = loadConfig()
    loginTwitter()
    googleCreds = loginGoogle()
    openSpreadsheet(googleCreds)
    exportMentions()
    exportDMs()
    saveConfig(config)


if __name__ == '__main__':
    main()
