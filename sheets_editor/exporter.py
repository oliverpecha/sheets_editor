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
                raise  # Re-raise the exception

    def export_table(
        self, data: List[Dict], version: str, sheet_name: str,
        columns: Optional[List[str]] = None,
        spreadsheet: Optional[gspread.Spreadsheet] = None,
        formatting = None,
        conditional_format = None
    ) -> None:
        """Exports data to a Google Sheet, handling existing sheets."""
        spreadsheet_name = f"{self.config.file_name}_{version}"
    
        if spreadsheet is None:
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
            worksheet = spreadsheet.worksheet(sheet_name)
            worksheet.clear()  # Clear the sheet if it exists
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(columns) if columns else 1)
    
        # Write data (handle empty data case)
        if columns:
            worksheet.append_row(columns)
        if data:
            rows = [[str(row.get(col, '')) for col in columns] for row in data]
            worksheet.append_rows(rows)
    
        #->>>>>> Only apply formatting if it is provided
        #print(f"Formatting value in export_table: {formatting}")  # Check the value BEFORE the if
        #if formatting is not None and formatting != {}:
            #self.formatter.format_worksheet(worksheet, formatting)

        if formatting or conditional_formats:
            print("Formatting about to be applied...")
            #self.formatter.format_worksheet(worksheet, formatting, conditional_formats)
            self.formatter.format_worksheet(worksheet, formatting, conditional_formats or []) # Pass empty list if None

        
    
        return spreadsheet.url
