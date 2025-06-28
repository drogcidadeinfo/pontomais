import os
import glob
import json
import logging
from datetime import datetime

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Constants & logging
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Environment variables
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
CREDENTIALS_JSON = os.getenv("PONTOMAIS_CRED")

if not SPREADSHEET_ID:
    raise ValueError("SPREADSHEET_ID environment variable not set.")
if not CREDENTIALS_JSON:
    raise ValueError("PONTOMAIS_CRED environment variable not set.")

# Helper functions

def get_latest_file(extension: str = 'xlsx', directory: str = '.') -> str | None:
    files = glob.glob(os.path.join(directory, f'*.{extension}'))
    if not files:
        logging.warning("No Excel files found.")
        return None
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def load_excel_file(path: str) -> pd.DataFrame | None:
    try:
        return pd.read_excel(path, skiprows=3)
    except Exception as e:
        logging.error(f"Failed to load Excel file: {e}")
        return None

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip()
    df = df.dropna(how='all')

    if 'Matrícula' in df.columns:
        df = df.drop(columns=['Matrícula'])

    if 'Nome' in df.columns:
        resumo_index = df[df['Nome'].astype(str).str.strip().str.lower() == 'resumo'].index
        if not resumo_index.empty:
            df = df.loc[:resumo_index[0] - 1]

    return df

def get_gspread_client() -> gspread.Client:
    creds_dict = json.loads(CREDENTIALS_JSON)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds), creds

def update_google_sheet_data(df: pd.DataFrame, sheet_name: str, spreadsheet_id: str) -> tuple[gspread.Worksheet | None, Credentials | None]:
    client, creds = get_gspread_client()
    try:
        sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
        sheet.clear()

        data = [df.columns.tolist()] + df.fillna("").values.tolist()
        sheet.update(range_name="A1", values=data)
        logging.info("✅ Data uploaded to Google Sheets successfully.")
        return sheet, creds
    except Exception as e:
        logging.error(f"Failed to update Google Sheets: {e}")
        return None, None

def apply_sheet_formatting(spreadsheet_id: str, sheet_id: int, creds: Credentials, num_columns: int) -> None:
    try:
        service = build('sheets', 'v4', credentials=creds)

        requests = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "textFormat": {"bold": True}
                        }
                    },
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment)"
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": num_columns
                    }
                }
            }
        ]

        body = {'requests': requests}
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()

        logging.info("✅ Formatting applied to Google Sheet.")
    except Exception as e:
        logging.error(f"Failed to apply formatting: {e}")

# Main flow

def main():
    file_path = get_latest_file(extension='xlsx', directory='/home/runner/work/pontomais/pontomais/')
    if not file_path:
        logging.info("No files found to process.")
        return

    logging.info(f"📄 Processing file: {file_path}")
    df = load_excel_file(file_path)
    if df is None:
        return

    df_clean = clean_dataframe(df)
    sheet, creds = update_google_sheet_data(df_clean, sheet_name='dados', spreadsheet_id=SPREADSHEET_ID)

    if sheet and creds:
        sheet_id = sheet._properties.get('sheetId')
        if sheet_id is not None:
            apply_sheet_formatting(SPREADSHEET_ID, sheet_id, creds, num_columns=len(df_clean.columns))
        else:
            logging.warning("Sheet ID could not be retrieved; skipping formatting.")

if __name__ == "__main__":
    main()
