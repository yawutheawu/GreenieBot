import discord
import os
from dotenv import load_dotenv
from pathlib import Path
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import google.auth

import SheetGetter

SheetsAPIFile = Path(r"Sensitive/Hidden/SheetsAPI.json")
environmentFile = Path(r"Sensitive/Hidden/Variables.env")
TokenPath = Path(r"Sensitive/Hidden/token.json")
load_dotenv(environmentFile)
BotAPI = os.getenv("DiscordAPIKey")
SheetLink = os.getenv("SheetLink")
if SheetLink is None and BotAPI is None:
    raise Exception(".env file with secrets not found")
SheetID = SheetLink.split("/")[7]
TrapSheetRange = "'Trap Sheet'!A:I"
GreenieBoardData = SheetGetter.getSheetContents(SheetID, 'Greenie Board')
TrapSheetData = SheetGetter.getRange(SheetID,TrapSheetRange)
#GreenieBoard dict struct: Greenie = 
#Trap Sheet dict struct: TrapSheet = {"Username in Sheet" : {Date _ Attempt : {Type : "", Date : "", Grade : "", Score : "", Wire : "", Comments : "", Grader : ""}}}
TrapSheet = {}
def populateTrapsheet(Data):
    print(Data)
populateTrapsheet(TrapSheetData)