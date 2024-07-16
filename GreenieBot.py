import discord
import os
from dotenv import load_dotenv
from pathlib import Path
environmentFile = Path(r"Sensitive/Hidden/Variables.env")
load_dotenv(environmentFile)
BotAPI = os.getenv("DiscordAPIKey")
SheetsAPI = os.getenv("SheetsAPIKey")
SheetLink = os.getenv("SheetLink")