from unittest.mock import MagicMock, patch
import pandas as pd
import pytest
from src.sheets_client import SheetsClient


@patch("src.sheets_client.gspread.authorize")
@patch("src.sheets_client.Credentials.from_service_account_file")
def make_client(mock_creds, mock_authorize):
    mock_spreadsheet = MagicMock()
    mock_authorize.return_value.open_by_key.return_value = mock_spreadsheet
    client = SheetsClient(spreadsheet_id="dummy_id", credentials_path="dummy.json")
    client._spreadsheet = mock_spreadsheet
    return client, mock_spreadsheet


def test_read_sheet_returns_dataframe():
    client, mock_spreadsheet = make_client()
    mock_ws = MagicMock()
    mock_ws.get_all_records.return_value = [{"name": "Alice", "age": 30}]
    mock_spreadsheet.worksheet.return_value = mock_ws

    df = client.read_sheet("Sheet1")

    assert isinstance(df, pd.DataFrame)
    assert list(df.columns) == ["name", "age"]
    assert df.iloc[0]["name"] == "Alice"


def test_list_sheets_returns_titles():
    client, mock_spreadsheet = make_client()
    mock_ws1 = MagicMock()
    mock_ws1.title = "Sheet1"
    mock_ws2 = MagicMock()
    mock_ws2.title = "Sheet2"
    mock_spreadsheet.worksheets.return_value = [mock_ws1, mock_ws2]

    result = client.list_sheets()

    assert result == ["Sheet1", "Sheet2"]
