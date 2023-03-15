from __future__ import print_function

import os.path #dient om te kijken of token al is gemaakt

from google.auth.transport.requests import Request  #dient om de user in te laten loggen
from google.oauth2.credentials import Credentials  #dient om creds uit token.json te halen als het bestaat
from google_auth_oauthlib.flow import InstalledAppFlow #dient om de juiste file uit credentials te halen
from googleapiclient.discovery import build #weet ik nog niet
from googleapiclient.errors import HttpError #print de error indien nodig

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1va20jTEx0PSwK6B5d24uQ5zycEWu5rghmSlIUQHRC-E'
SAMPLE_RANGE_NAME = 'A:NG'
TE_SCHRIJVEN_RANGE = 'F:NH'

#1. TOKEN DEFINIEREN
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

#2. PROGRAMMA STARTEN


#KIJKEN OF DE WAARDE VERANDERD IS


# DE WAARDE DOORGEVEN
# Call the Sheets API
import requests
import json
response = requests.get(f'https://sheets.googleapis.com/v4/spreadsheets/{SAMPLE_SPREADSHEET_ID}/values/{SAMPLE_RANGE_NAME}', headers={'Authorization': f'Bearer {creds.token}'})

# Check if the request was successful
if response.status_code == 200:
    # Retrieve the values from the response
    data = response.json()
    values = data.get('values', [])

    NieuwResponse = requests.get(f'https://sheets.googleapis.com/v4/spreadsheets/{SAMPLE_SPREADSHEET_ID}/values/{TE_SCHRIJVEN_RANGE}', headers={'Authorization': f'Bearer {creds.token}'})
    NieuwData = NieuwResponse.json()
    NieuwValues = NieuwData.get('values', [])
     

    # Loop through each row of the range and compare column A to column NG
    for i, row in enumerate(values[1:], start=2):  # Skip the first row, start at row 2
        print(i)
        if row[0] != row[-1]:
            # Update column NG with the value of column A
            Vervang = row[0]
            ggids = row[2]
            gids = int(ggids[-5:])
            print(gids)
            nieuw = NieuwValues[i-1]
            print(nieuw)


            # Build the update request
            update_range = f'NG{i}'
            update_data = {                
                'range': update_range,
                'majorDimension': 'ROWS',
                'values': [[Vervang]]
            }
            
 

            # VERVANGEN IN NINOX
            field_id = "AVVAIL"
            loginx = "55af8410-b2c9-11ed-b17d-67f15f6a0397"


            # Checken of de code juist is
            # URL voor het ophalen van de record die moet worden gewijzigd


            # URL voor het ophalen van de record die moet worden gecontroleerd en gewijzigd
            url = f"https://api.ninox.com/v1/teams/w4tgKLQYBYpNiJux9/databases/yl815wodrqbw/tables/Contacten/records/{gids}"

            # Headers voor de API-aanroep
            headers = {
                "Authorization": f"Bearer {loginx}",
                "Content-Type": "application/json"
            }


            antwoord = requests.get(url, headers=headers)
            DataAntwoord = antwoord.json()
            print(DataAntwoord)
            if 'Code' in DataAntwoord['fields']:
                code_value = DataAntwoord['fields']['Code']
 
                print(code_value)
                print(ggids)

                if code_value == ggids:
                    

                    # Nieuwe inhoud van het veld "field"
                    new_field_value = nieuw

                    # JSON-payload voor de PUT-aanroep
                    payload = {
                        "fields": {
                            field_id: new_field_value
                        }
                    }

                    # PUT-aanroep naar de Ninox REST API om het veld "field" te wijzigen
                    response = requests.put(url, headers=headers, data=json.dumps(payload))

                    # Controleer of de PUT-aanroep succesvol was
                    if response.status_code == 200:
                        print("Ninox succesvol bijgewerkt!")

                        # Send the update request
                        update_response = requests.put(f'https://sheets.googleapis.com/v4/spreadsheets/{SAMPLE_SPREADSHEET_ID}/values/{update_range}?valueInputOption=USER_ENTERED', headers={'Authorization': f'Bearer {creds.token}'}, json = update_data)


                        # Check if the update was successful
                        if update_response.status_code == 200:
                            print(f'Row {values.index(row)+1} updated successfully')
                        else:
                            print(f'Error updating row {values.index(row)+1}')
                            print(f'Error updating row {row}: {update_response.text}')

                    else:
                        print(f"Fout bij het bijwerken van het veld: {response.status_code} - {response.text}")

                else:
                    print("VERKEERDE CODE")
            else:
                print("GEEN CODE")

else:
    print('Error retrieving values')


