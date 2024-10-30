from google.oauth2.service_account import Credentials
import gspread
from typing import Dict, List, Optional, Any
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

    def export_table(self, 
                    data: List[Dict],
                    version: str, 
                    sheet_name: str,
                    columns: Optional[List[str]] = None) -> None:
        """
        Main export method
        
        Args:
            data: List of dictionaries containing the data
            version: Version identifier for the sheet
            sheet_name: Name of the worksheet
            columns: Optional list of columns to include
        """
        # Determine columns if not provided
        if not columns and data:
            columns = list(data[0].keys())
        
        # Remove ignored columns
        if self.config.ignore_columns:
            columns = [c for c in columns if c not in self.config.ignore_columns]

        # Create spreadsheet name
        spreadsheet_name = f"{self.config.file_name}_{version}"
        
        creds = Credentials.from_service_account_info(self.credentials, scopes=self.scope)
        client = gspread.authorize(creds)

        try:
            # Create or open spreadsheet
            try:
                spreadsheet = client.open(spreadsheet_name)
            except gspread.SpreadsheetNotFound:
                spreadsheet = client.create(spreadsheet_name)
                
            # Create or get worksheet
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
            except gspread.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(sheet_name, 1000, len(columns))

            # Write data
            worksheet.clear()
            worksheet.append_row(columns)
            
            # Convert data to rows
            rows = [[str(row.get(col, '')) for col in columns] for row in data]
            if rows:
                worksheet.append_rows(rows)

            # Apply formatting
            formatting = {
                'alternate_rows': True,
                'row_height': 42,
                'background_color': self.config.alternate_row_color
            }
            self.formatter.format_worksheet(worksheet, formatting)

            # Share if needed
            if self.config.share_with:
                for email in self.config.share_with:
                    spreadsheet.share(email, perm_type='user', role='writer')

            print(f"Data exported to: {spreadsheet.url}")
            return spreadsheet.url

        except Exception as e:
            print(f"Error in sheet export: {e}")
            raise