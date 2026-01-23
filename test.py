from googleapiclient.discovery import build
from src.Auth.auth_sheets import autenticar

SPREADSHEET_ID = '1qFfGUnS_h86kV91MpTN0EyrZhS0-uAoKpZSpIy_1HsA'
RANGE = 'A1:CU5'

creds = autenticar()
service = build('sheets', 'v4', credentials=creds)

result = service.spreadsheets().values().get(
    spreadsheetId=SPREADSHEET_ID,
    range=RANGE,
).execute()

values = result.get('values', [])

for linha in values:
    print(linha)
