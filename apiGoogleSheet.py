from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os
import csv

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def connection_to_sheet():
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds

def is_downloaded_csv_google_sheet(credentials, spreadsheet_id, spreadsheet_range):
    if (credentials and spreadsheet_id and spreadsheet_range):
        try:
            service = build('sheets', 'v4', credentials=credentials)

            sheet = service.spreadsheets()
            spreadsheets = sheet.values().get(spreadsheetId=spreadsheet_id, range=spreadsheet_range).execute()

            convert_spreadsheet_to_csv(spreadsheets)

            print('Téléhcargement du fichier google sheet terminé')
            return True
        except HttpError as err:
            print(err)
            return False
    else:
        print(f'credentials: {credentials}, spreadsheet_id: {spreadsheet_id} et spreadsheet_range {spreadsheet_range} sont invalide')
        return False
        
def convert_spreadsheet_to_csv(spreadsheet):
    csv_file = 'data.csv'

    if os.path.exists(csv_file):
        os.remove(csv_file)

    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        data = spreadsheet.get('values')
        for row in data:
            try:
               price = row[4] 
               price = float(price.replace(',', '.'))
               row[4] = price
            except:
                pass
            writer.writerow(row)
    f.close()


