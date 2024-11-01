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
        gc = gspread.authorize(Credentials.from_service_account_info(self.credentials, scopes=self.scope))
        try:
            spreadsheet = gc.open(spreadsheet_name)
            print(f"Opened existing spreadsheet: {spreadsheet_name}")
            return spreadsheet
        except gspread.SpreadsheetNotFound:
            try:
                spreadsheet = gc.create(spreadsheet_name)
                print(f"Created new spreadsheet: {spreadsheet_name}")
                if self.config.share_with:
                    for email in self.config.share_with:
                        spreadsheet.share(email, perm_type='user', role='writer')
                return spreadsheet
            except gspread.exceptions.APIError as e:
                print(f"Error creating spreadsheet: {e}")
                raise  # Or handle the error as needed

    def export_table(self, 
                    data: List[Dict],
                    version: str, 
                    sheet_name: str,
                    columns: Optional[List[str]] = None) -> None:
        """Exports data to a Google Sheet, overwriting Sheet1 if it's the only sheet."""
        spreadsheet_name = f"{self.config.file_name}_{version}"
        spreadsheet = self._open_spreadsheet(spreadsheet_name)

        if spreadsheet is None:
            raise ValueError("Could not open or create spreadsheet")

        # Determine columns if not provided
        if not columns and data:
            columns = list(data[0].keys())

        # Remove ignored columns
        if self.config.ignore_columns:
            columns = [c for c in columns if c not in self.config.ignore_columns]

        try:
            worksheet = spreadsheet.worksheet("Sheet1") # Try to get Sheet1
            if spreadsheet.worksheets() is not None and len(spreadsheet.worksheets()) == 1:
                worksheet.update_title(sheet_name)  # Rename Sheet1 if it's the only sheet
            else:
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(columns) if columns else 1)  # Explicitly naming parameters for clarity
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(columns) if columns else 1)  # Create if it doesn't exist


        # Write data (handling the case where data might be empty)
        worksheet.clear()
        if columns:  # Only append header row if there are columns
            worksheet.append_row(columns)
        if data:  # Only append data rows if data is not empty
            rows = [[str(row.get(col, '')) for col in columns] for row in data]
            worksheet.append_rows(rows)


        # Apply formatting
        formatting = {
            'alternate_rows': True,
            'row_height': 42,
            'background_color': self.config.alternate_row_color
        }
        self.formatter.format_worksheet(worksheet, formatting)

        return spreadsheet.url # Return the spreadsheet URL
