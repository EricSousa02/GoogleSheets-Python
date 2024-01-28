# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
# https://developers.google.com/sheets/api/quickstart/python?hl=pt-br
# python main.py
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_ID = 'SPREADSHEET_ID'
RANGE_NAME = 'A4:H27'

def calculate_status(row):
    # Calculate the average of P1, P2, and P3
    average = round((row[1] + row[2] + row[3]) / 30, 1)

    # Check if the student is Reprovado por Falta
    if row[0] > 0.25 * 60:
        return 'Reprovado por Falta'

    # Determine the student's situation based on the average
    if average < 5:
        return 'Reprovado por Nota'
    elif 5 <= average < 7:
        return 'Exame Final'
    else:
        return 'Aprovado'

def calculate_naf(row):
    # Calculate the Nota para Aprovacao Final (naf) for Exame Final
    if row[3] == 'Exame Final':
        average = round((row[0] + row[1] + row[2]) / 30, 1)
        naf = 2 * (7 - average)
        return round(naf, 1)

    # If not Exame Final, set Nota para Aprovacao Final to 0
    return 0

def main():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        
        # Get values from the spreadsheet
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
        values = result.get('values', [])
        


        if not values:
            print('No data found.')
        else:
            for row in values:
             # Calculate student status considering only columns 3 to 6
                if len(row) >= 7 and row[6]:  
                    row[6] = (calculate_status([float(cell) for cell in row[2:6]]))    
                else:
                    row.append(calculate_status([float(cell) for cell in row[2:6]]))
                    
                      
            # Add a new column for Nota para Aprovação Final (naf)
            for row in values:
                if len(row) >= 8 and row[7]:
                    converted_values = [float(cell) if index != 3 else cell for index, cell in enumerate(row[3:7])]
                    row[7] = (calculate_naf(converted_values))          
                else:    
                    converted_values = [float(cell) if index != 3 else cell for index, cell in enumerate(row[3:7])]
                    row.append(calculate_naf(converted_values))

            sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, body={'values': values}, valueInputOption='USER_ENTERED').execute()

            print('completed.')

    except HttpError as err:
        print(err)

if __name__ == "__main__":
    main()
