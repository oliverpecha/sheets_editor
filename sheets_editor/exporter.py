import gspread
from typing import Dict, List, Optional, Any
from .config import SheetConfig
from .formatter import SheetFormatter
from google.oauth2.service_account import Credentials

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
        gc = gspread.authorize(Credentials.from_service_account_info(self.credentials, scopes=self.scope)) #scopes added
        try:
            spreadsheet = gc.open(spreadsheet_name)
            print(f"Opened existing spreadsheet: {spreadsheet_name}")
            return spreadsheet
        except gspread.SpreadsheetNotFound:
            try:
                spreadsheet = gc.create(spreadsheet_name)
                print(f"Created new spreadsheet: {spreadsheet_name}")
                if self.config.share_with: #Share during creation
                    for email in self.config.share_with:
                        spreadsheet.share(email, perm_type='user', role='writer')
                return spreadsheet
            except gspread.exceptions.APIError as e:
                print(f"Error creating spreadsheet: {e}")
                raise  # Or handle the error as needed (e.g., return None)

    def export_table(self, 
                    data: List[Dict],
                    version: str, 
                    sheet_name: str,
                    columns: Optional[List[str]] = None,
                    delete_sheet1: bool = True) -> None:
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

            # Create or get worksheet
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
            except gspread.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(sheet_name, 1000, len(columns))

            # Write data
            worksheet.clear()  # Clear existing data
            worksheet.append_row(columns) # Append header row
            rows = [[str(row.get(col, '')) for col in columns] for row in data] # Create data rows
            if rows:
                worksheet.append_rows(rows) #Append data rows

            # Apply formatting
            formatting = {
                'alternate_rows': True,
                'row_height': 42, #Make this configurable later if needed
                'background_color': self.config.alternate_row_color
            }
            self.formatter.format_worksheet(worksheet, formatting)

            if delete_sheet1:
                self._delete_empty_sheet1(spreadsheet)

            print(f"Data exported to: {spreadsheet.url}")
            return spreadsheet.url

        except Exception as e:
            print(f"Error in sheet export: {e}")
            raise
            
    def _delete_empty_sheet1(self, spreadsheet: gspread.Spreadsheet) -> None:
        """Deletes Sheet1 if it's empty and not the only sheet."""
        if spreadsheet.worksheets() is not None and len(spreadsheet.worksheets()) > 1:
            try:
                sheet1 = spreadsheet.worksheet("Sheet1")
                if not sheet1.get_all_values():
                    spreadsheet.del_worksheet(sheet1)
                    print("Successfully deleted Sheet1") #Keep this here for explicit logging
            except gspread.exceptions.WorksheetNotFound:
                print("Sheet1 not found")
        else:
            print("Not deleting Sheet1. Either only sheet or API issue.")
