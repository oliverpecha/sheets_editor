from typing import Any, Dict, List

class SheetFormatter:
    @staticmethod
    def format_worksheet(worksheet: Any, formatting_config: Dict) -> None:
        """Applies formatting to the worksheet based on the formatting_config."""
        try:
            print(f"Worksheet object: {worksheet}")
            print("Formatting dictionary:", formatting_config)
            print("Bold rows config:", formatting_config.get('bold_rows'))

            values = worksheet.get_all_values()
            if not values:  # Handle empty sheet
                return

            sheet_id = worksheet._properties['sheetId']
            num_rows = len(values)
            num_cols = len(values[0])

            requests = []

            # Alternate Row Formatting
            if formatting_config.get('alternate_rows'):
                for row in range(2, num_rows + 1, 2):  # Start from row 2, skip header
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


            # Bold Row Formatting
            if 'bold_rows' in formatting_config:
                bold_rows = formatting_config['bold_rows']
                print("Bolding rows:", bold_rows)
                for row in bold_rows:
                    print("Bolding row:", row)
                    requests.append({
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": row - 1,  # gspread indices are 0-based
                                "endRowIndex": row,
                                "startColumnIndex": 0,
                                "endColumnIndex": num_cols
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "bold": True
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.bold"
                        }
                    })

            # Batch Update (execute all formatting requests)
            if requests:
                worksheet.spreadsheet.batch_update({"requests": requests})

        except Exception as e:
            print(f"Error in formatting: {e}")
            raise  # Reraise for debugging
