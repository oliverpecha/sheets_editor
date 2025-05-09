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
        self, data: List[Dict], sheet_name: str, version: Optional[str] = None,
        columns: Optional[List[str]] = None,
        spreadsheet: Optional[gspread.Spreadsheet] = None,
        formatting: Optional[Dict] = None,
        conditional_formats: Optional[List[Dict]] = None
    ) -> None:
        """Exports data to a Google Sheet, handling existing sheets and optional versions."""
    
        # Check if data is empty and handle it gracefully
        if not data:
            print(f"⚠️ No data to export to sheet '{sheet_name}'. Skipping export process.")
            return None
    
        if version:  # Append version if provided
            spreadsheet_name = f"{self.config.file_name}_{version}"
        else:
            spreadsheet_name = self.config.file_name  # Use the base file name
    
        if spreadsheet is None:
            spreadsheet = self._open_spreadsheet(spreadsheet_name)
        
        if spreadsheet is None:
            raise ValueError("Could not open or create spreadsheet")
    
        # Determine columns if not provided
        if not columns and data:
            columns = list(data[0].keys())
        
        # Ensure columns is not empty
        if not columns:
            print(f"⚠️ No columns defined for export to sheet '{sheet_name}'. Creating a default column.")
            columns = ["empty_data"]  # Use a default column name
    
        # Remove ignored columns
        if self.config.ignore_columns:
            columns = [c for c in columns if c not in self.config.ignore_columns]
    
        # Check if worksheet exists without raising an exception
        worksheet = None
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
            worksheet.clear()  # Clear existing sheet if it exists
            print(f"Cleared existing sheet: {sheet_name}")
        except gspread.WorksheetNotFound:
            # Suppress the error output - just create a new worksheet
            print(f"Sheet '{sheet_name}' not found. Creating a new one.")
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(columns))
    
        # Write headers and data
        if columns:
            worksheet.append_row(columns)
        if data:
            rows = [[str(row.get(col, '')) for col in columns] for row in data]
            worksheet.append_rows(rows)
    
        # Apply formatting if provided
        if formatting or conditional_formats: # 'formatting' here is your {(r,c):style} dict
            print("Applying formatting...")
            self.formatter.format_worksheet(
                worksheet=worksheet,
                # general_rules_config=None, # Or pass a different dict if you have general rules
                targeted_cell_formats=formatting, # This is your {(row,col):style} dict
                conditional_formats=conditional_formats or []
            )

         
        return spreadsheet.url
