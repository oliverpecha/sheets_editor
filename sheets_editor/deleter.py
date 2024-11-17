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


    def _ensure_default_sheet(self, spreadsheet):
        """Helper function to ensure a default sheet "Sheet1" exists."""
        default_sheet_name = "Sheet1"
        try:
            spreadsheet.worksheet(default_sheet_name)
            print(f"Default sheet '{default_sheet_name}' already exists.")
        except gspread.WorksheetNotFound:
            spreadsheet.add_worksheet(title=default_sheet_name, rows=100, cols=20)
            print(f"Added default sheet: {default_sheet_name}")


    def delete_single_sheet(self, spreadsheet_name, sheet_name):
        """Deletes a specific sheet, ensuring it's not the only one."""
        spreadsheet = self._open_spreadsheet(spreadsheet_name)
        if spreadsheet:
            try:
                worksheets = spreadsheet.worksheets()

                if len(worksheets) == 1 and worksheets[0].title == sheet_name:
                    # It's the only sheet, so create the default sheet first
                    self._ensure_default_sheet(spreadsheet)

                # Now delete the specified sheet (if it exists after ensuring default)
                try:
                    sheet_to_delete = spreadsheet.worksheet(sheet_name)  # Raises exception if not found
                    print(f"Deleting sheet: {sheet_to_delete.title}")
                    spreadsheet.del_worksheet(sheet_to_delete)
                    print(f"Deleted sheet: {sheet_name}")

                except gspread.WorksheetNotFound:
                    print(f"Sheet '{sheet_name}' not found.")

            except Exception as e:
                print(f"An error occurred while deleting the sheet: {e}")



    def delete_all_sheets(self, spreadsheet_name):
        """Ensures a default sheet 'Sheet1' exists, then deletes all other sheets."""
        spreadsheet = self._open_spreadsheet(spreadsheet_name)
        if spreadsheet:
            try:

                # 1. Ensure the default sheet exists (using helper function)
                self._ensure_default_sheet(spreadsheet)

                # 2. Get the updated list of worksheets
                worksheets = spreadsheet.worksheets()
                default_sheet_name = "Sheet1"  # Assuming Sheet1 is the default name

                # 3. Delete all sheets that are not the default sheet
                for worksheet in worksheets:
                    if worksheet.title != default_sheet_name:
                        print(f"Deleting sheet: {worksheet.title}")
                        spreadsheet.del_worksheet(worksheet)
                        print(f"Deleted sheet: {worksheet.title}")

                print("All non-default sheets deleted.")

            except Exception as e:
                print(f"An error occurred while deleting/adding sheets: {e}")
