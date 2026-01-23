from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os.path

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def autenticar():
    creds = None

    if os.path.exists('src/auth/token.json'):
        creds = Credentials.from_authorized_user_file('src/auth/token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'src/auth/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('src/auth/token.json', 'w') as token:
            token.write(creds.to_json())

    return creds
