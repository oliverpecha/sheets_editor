import gspread
from typing import Dict, List, Optional, Any
from gspread.utils import ValueInputOption
from google.oauth2.service_account import Credentials
from .config import SheetConfig
from .formatter import SheetFormatter

class SheetsExporter:
    def __init__(self, credentials: Dict, config: SheetConfig):
        self.credentials = credentials
        self.config = config
        self.scope = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        self.formatter = SheetFormatter()

    def _open_spreadsheet(self, spreadsheet_name: str):
        """Opens or creates a spreadsheet, handling potential errors."""
        gc = gspread.authorize(Credentials.from_service_account_info(self.credentials, scopes=self.scope))
        try:
            spreadsheet = gc.open(spreadsheet_name)
            print(f"Opened existing spreadsheet: {spreadsheet_name}")
        except gspread.SpreadsheetNotFound:
            try:
                spreadsheet = gc.create(spreadsheet_name)
                print(f"Created new spreadsheet: {spreadsheet_name}")
                if self.config.share_with:
                    for email in self.config.share_with:
                        spreadsheet.share(email, perm_type='user', role='writer')
            except gspread.exceptions.APIError as e:
                print(f"Error creating spreadsheet: {e}")
                raise
        return spreadsheet

    def export_table(self, 
                     data: List[Dict],
                     version: str, 
                     sheet_name: str,
                     columns: Optional[List[str]] = None,
                     delete_sheet1: bool = True) -> str:
        """Exports data to a Google Sheet."""
        spreadsheet_name = f"{self.config.file_name}_{version}"
        try:
            spreadsheet = self._open_spreadsheet(spreadsheet_name)
            if spreadsheet is None:
                raise ValueError("Could not open or create spreadsheet")

            # Determine columns if not provided
            if not columns and data:
                columns = list(data[0].keys())

            # Remove ignored columns
            if self.config.ignore_columns:
                columns = [c for c in columns if c not in self.config.ignore_columns]

            # Try to get the desired worksheet
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
                print(f"Found existing worksheet: {sheet_name}")
            except gspread.WorksheetNotFound:
                print(f"Worksheet '{sheet_name}' not found.")
                # Check if 'Sheet1' exists and is empty
                try:
                    sheet1 = spreadsheet.worksheet("Sheet1")
                    sheet1_values = sheet1.get_all_values()
                    if not sheet1_values or sheet1_values == [['']]:
                        # 'Sheet1' is empty, rename and use it
                        sheet1.update_title(sheet_name)
                        worksheet = sheet1
                        print(f"Renamed 'Sheet1' to '{sheet_name}' and reusing it.")
                    else:
                        # 'Sheet1' has data, create a new worksheet
                        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols=str(len(columns)))
                        print(f"Created new worksheet: {sheet_name}")
                except gspread.WorksheetNotFound:
                    # 'Sheet1' does not exist, create a new worksheet
                    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols=str(len(columns)))
                    print(f"Created new worksheet: {sheet_name}")

            # Write data to the worksheet
            worksheet.clear()  # Clear existing data
            worksheet.append_row(columns)  # Append header row
            rows = [[str(row.get(col, '')) for col in columns] for row in data]  # Create data rows
            if rows:
                worksheet.append_rows(rows, value_input_option=ValueInputOption.raw)  # Append data rows

            # Apply formatting
            formatting = {
                'alternate_rows': True,
                'row_height': 42,  # Make this configurable later if needed
                'background_color': self.config.alternate_row_color
            }
            self.formatter.format_worksheet(worksheet, formatting)

            print(f"Data exported successfully to worksheet: {sheet_name}")

            # Optionally delete 'Sheet1' if it's empty and not the only sheet
            if delete_sheet1:
                self._delete_empty_sheet1(spreadsheet)

            return spreadsheet.url

        except Exception as e:
            print(f"Error in sheet export: {e}")
            raise

    def _delete_empty_sheet1(self, spreadsheet: gspread.Spreadsheet) -> None:
        """Deletes 'Sheet1' if it's empty and not the only sheet."""
        worksheets = spreadsheet.worksheets()
        if len(worksheets) > 1:
            try:
                sheet1 = spreadsheet.worksheet("Sheet1")
                sheet1_values = sheet1.get_all_values()
                if not sheet1_values or sheet1_values == [['']]:
                    spreadsheet.del_worksheet(sheet1)
                    print("Deleted empty 'Sheet1'.")
                else:
                    print("'Sheet1' is not empty, not deleting.")
            except gspread.WorksheetNotFound:
                print("'Sheet1' not found, nothing to delete.")
        else:
            print("Cannot delete 'Sheet1'. It is the only worksheet in the spreadsheet.")
