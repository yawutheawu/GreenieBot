import os
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import google.auth

#Sheets API Quickstart: https://developers.google.com/drive/api/quickstart/python#step_1_turn_on_the
#Get Whole Sheet: https://stackoverflow.com/questions/51851389/google-sheet-python-api-how-get-values-in-first-sheet-if-dont-know-range
#On Edit: https://developers.google.com/drive/api/guides/manage-changes#python, https://developers.google.com/drive/api/guides/push
#Get Color of cell: https://stackoverflow.com/questions/41981911/how-to-read-the-color-of-a-cell-in-google-sheets


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

def directToLetter(Num):
    match Num:
        case 1:
            return 'A'
        case 2:
            return 'B'
        case 3:
            return 'C'
        case 4:
            return 'D'
        case 5:
            return 'E'
        case 6:
            return 'F'
        case 7:
            return 'G'
        case 8:
            return 'H'
        case 9:
            return 'I'
        case 10:
            return 'J'
        case 11:
            return 'K'
        case 12:
            return 'L'
        case 13:
            return 'M'
        case 14:
            return 'N'
        case 15:
            return 'O'
        case 16:
            return 'P'
        case 17:
            return 'Q'
        case 18:
            return 'R'
        case 19:
            return 'S'
        case 20:
            return 'T'
        case 21:
            return 'U'
        case 22:
            return 'V'
        case 23:
            return 'W'
        case 24:
            return 'X'
        case 25:
            return 'Y'
        case 26:
            return 'Z'
        case _:
            return ""

     
def convertToLetter(n):
    numCol = [0,0]
    divided = n
    result = ""
    while divided >= 1:
        numCol[-1] += 1
        for i in range(len(numCol)-1,0,-1):
            if numCol[i]>26:
                numCol[i] = numCol[i] - 26
                numCol[i-1] +=1
        if numCol[0] > 26:
            numCol[0] = numCol[0]-26
            numCol.insert(0,1)
        divided-=1
    for i in numCol:
        result += directToLetter(i)
    return result

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

def getSheetContents(SheetID,SheetName):
    fetchedRange = ""
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        SheetChoices = sheet.get(spreadsheetId=SheetID, fields='sheets(data/rowData/values/userEnteredValue,properties(index,sheetId,title))').execute()
        appropIndex = [0,0]
        sheetFound = False
        for i in SheetChoices["sheets"]:
           if i["properties"]['title'] == SheetName:
               appropIndex = [i["properties"]['index'],i["properties"]['sheetId']]
               sheetFound = True
        if sheetFound:
            sheetName = SheetChoices['sheets'][appropIndex[0]]['properties']['title']
            lastColumn = convertToLetter(max([len(e['values']) for e in SheetChoices['sheets'][appropIndex[0]]['data'][0]['rowData'] if e]))
            result = (
                sheet.values()
                .get(spreadsheetId=SheetID, range=str(str(sheetName)+"!"+"A:"+str(lastColumn)))
                .execute()
                )
            values = result.get("values", [])
            if not values:
                print("No data found.")
                return
            return values
        else:
            print("Sheet Not Found")
            return
    except Exception as err:
        print(err)
        return