from typing import Any, Dict

class SheetFormatter:
    @staticmethod
    def format_worksheet(worksheet: Any, formatting_config: Dict) -> None:
        print("Formatting dictionary:", formatting_config)
        """Apply formatting to worksheet"""
        try:
            print(f"Worksheet object: {worksheet}") # Check the worksheet object
            print("Formatting dictionary:", formatting_config)
            values = worksheet.get_all_values()
            if not values:
                return
                
            sheet_id = worksheet._properties['sheetId']
            num_rows = len(values)
            num_cols = len(values[0])
            
            requests = []
            
            # Format even rows
            for row in range(2, num_rows + 1):  # Start from 2 to skip header
                if row % 2 == 0:
                    if formatting_config.get('row_height'):
                        requests.append({
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "dimension": "ROWS",
                                    "startIndex": row - 1,
                                    "endIndex": row
                                },
                                "properties": {
                                    "pixelSize": formatting_config['row_height']
                                },
                                "fields": "pixelSize"
                            }
                        })
                    
                    if formatting_config.get('background_color'):
                        requests.append({
                            "repeatCell": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": row - 1,
                                    "endRowIndex": row,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": num_cols
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "backgroundColor": formatting_config['background_color']
                                    }
                                },
                                "fields": "userEnteredFormat.backgroundColor"
                            }
                        })
            
            if requests:
                worksheet.spreadsheet.batch_update({"requests": requests})
                
        except Exception as e:
            print(f"My unexpected error occurred in format_worksheet: {e}")
            print(f"Error in formatting: {e}")
            raise
