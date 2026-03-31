import os
from dotenv import load_dotenv

load_dotenv()

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
CREDENTIALS_PATH = os.getenv("CREDENTIALS_PATH", "credentials/service_account.json")
