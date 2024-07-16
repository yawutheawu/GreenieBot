import os
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SheetsAPIFile = Path(r"Sensitive/Hidden/SheetsAPI.json")
TokenPath = Path(r"Sensitive/Hidden/token.json")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = None
if os.path.exists(TokenPath):
    creds = Credentials.from_authorized_user_file(TokenPath, SCOPES)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
else:
    flow = InstalledAppFlow.from_client_secrets_file(
        SheetsAPIFile, SCOPES
    )
    creds = flow.run_local_server(port=0)
    with open(TokenPath, "w") as token:
        token.write(creds.to_json())



def getRange(SheetID,RequiredRange):
    try:
       service = build("sheets", "v4", credentials=creds)
       sheet = service.spreadsheets()
       result = (
          sheet.values()
          .get(spreadsheetId=SheetID, range=RequiredRange)
          .execute()
          )
       values = result.get("values", [])
       if not values:
          print("No data found.")
          return
       return values
    except Exception as err:
        print(err)