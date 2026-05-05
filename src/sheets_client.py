import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from src.config import SPREADSHEET_ID, CREDENTIALS_PATH

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


class SheetsClient:
    def __init__(self, spreadsheet_id: str = None, credentials_path: str = None):
        self.spreadsheet_id = spreadsheet_id or SPREADSHEET_ID
        creds_path = credentials_path or CREDENTIALS_PATH

        creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        self._client = gspread.authorize(creds)
        self._spreadsheet = self._client.open_by_key(self.spreadsheet_id)

    def read_sheet(self, sheet_name: str) -> pd.DataFrame:
        """シート全体をDataFrameとして取得する"""
        worksheet = self._spreadsheet.worksheet(sheet_name)
        records = worksheet.get_all_records()
        return pd.DataFrame(records)

    def read_range(self, sheet_name: str, range_notation: str) -> pd.DataFrame:
        """A1記法で範囲を指定してDataFrameとして取得する（1行目をヘッダーとして使用）"""
        worksheet = self._spreadsheet.worksheet(sheet_name)
        values = worksheet.get(range_notation)
        if not values:
            return pd.DataFrame()
        headers = values[0]
        rows = values[1:]
        return pd.DataFrame(rows, columns=headers)

    def list_sheets(self) -> list[str]:
        """スプレッドシート内の全シート名を返す"""
        return [ws.title for ws in self._spreadsheet.worksheets()]
