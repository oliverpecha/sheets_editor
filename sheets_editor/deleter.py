import gspread

class SheetDeleter:
    def __init__(self, credentials: Dict):
        self.credentials = credentials
        self.scope = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        #self.gc = gspread.authorize(self.credentials)
        self.deleter = SheetDeleter()

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
        """Deletes all sheets in the spreadsheet."""
        spreadsheet = self._open_spreadsheet(spreadsheet_name)
        if spreadsheet:
            try:
                worksheets = spreadsheet.worksheets()
                for worksheet in worksheets:
                    print(f"Deleting sheet: {worksheet.title}")
                    spreadsheet.del_worksheet(worksheet)
                    print(f"Deleted sheet: {worksheet.title}")
            except Exception as e:
                print(f"An error occurred while deleting sheets: {e}")
