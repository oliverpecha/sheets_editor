import gspread
from typing import Dict
from google.oauth2.service_account import Credentials

class SheetDeleter:
    def __init__(self, credentials: Dict):
        self.credentials = credentials
        self.scope = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
    def _open_spreadsheet(self, spreadsheet_name):
        """Opens a spreadsheet."""
        gc = gspread.authorize(Credentials.from_service_account_info(self.credentials, scopes=self.scope))
        try:
            spreadsheet = gc.open(spreadsheet_name)
            print(f"Opened existing spreadsheet: {spreadsheet_name}")
            return spreadsheet
        except gspread.SpreadsheetNotFound:
            print(f"Spreadsheet '{spreadsheet_name}' not found.")
            return None

    def delete_default_sheet(self, spreadsheet_name):
        """Deletes the default 'Sheet1' if it exists in the spreadsheet."""
        spreadsheet = self._open_spreadsheet(spreadsheet_name)
        if spreadsheet:
            try:
                sheet1 = spreadsheet.worksheet('Sheet1')
                spreadsheet.del_worksheet(sheet1)
                print("Sheet1 deleted.")
            except gspread.WorksheetNotFound:
                print("Sheet1 didn't exist.")

    def delete_all_sheets(self, spreadsheet_name):
        """Deletes all sheets and then adds a mandatory default new sheet."""
        spreadsheet = self._open_spreadsheet(spreadsheet_name)
        if spreadsheet:
            try:
                worksheets = spreadsheet.worksheets()
                # Delete existing sheets
                for worksheet in worksheets:
                    print(f"Deleting sheet: {worksheet.title}")
                    spreadsheet.del_worksheet(worksheet)
                    print(f"Deleted sheet: {worksheet.title}")
    
                # Add a new sheet
                new_sheet_name = "Sheet1"
                spreadsheet.add_worksheet(title=new_sheet_name, rows=100, cols=20)
                print(f"Added new sheet: {new_sheet_name}")
    
            except Exception as e:
                print(f"An error occurred while deleting/adding sheets: {e}")
