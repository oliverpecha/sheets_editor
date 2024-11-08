from typing import Any, Dict, List

class SheetFormatter:
    def __init__(self):
      pass #Needed if no instance attributes
        
    def _create_request(self, row_index, num_cols, sheet_id, format_style, entire_row, col_index):
        start_col = 0 if entire_row else col_index
        end_col = num_cols if entire_row else col_index + 1

        user_entered_format = {}

        if isinstance(format_style, dict):
            if all(k in format_style for k in ("red", "green", "blue")):
                user_entered_format['backgroundColor'] = format_style
            else:
                user_entered_format.update(format_style)

        return {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_index,
                    "endRowIndex": row_index + 1,
                    "startColumnIndex": start_col,
                    "endColumnIndex": end_col
                },
                "cell": {
                    "userEnteredFormat": user_entered_format
                },
                "fields": "userEnteredFormat"  # Or specify more detailed fields if needed
            }
        }


    def format_worksheet(self, worksheet, formatting_config=None, conditional_formats=None):
        """Applies formatting to the worksheet."""

        if not formatting_config and not conditional_formats:
            return

        try:
            values = worksheet.get_all_values()
            if not values:
                return

            sheet_id = worksheet._properties['sheetId']
            num_rows = len(values)
            num_cols = len(values[0])
            requests = []

            if formatting_config:
                requests.extend(self._apply_absolute_formatting(formatting_config, sheet_id, num_rows, num_cols, values))

            if conditional_formats:
                requests.extend(self._apply_conditional_formatting(conditional_formats, sheet_id, num_cols, values))

            if requests:
                worksheet.spreadsheet.batch_update({"requests": requests})

        except Exception as e:
            print(f"Error in formatting: {e}")
            raise
            
    


            # Batch Update (execute all requests)
           ''' if requests:
                worksheet.spreadsheet.batch_update({"requests": requests})
'''

    def _apply_absolute_formatting(self, formatting_config, sheet_id, num_rows, num_cols, values):
        requests = []
        if formatting_config.get('alternate_rows'):
            for row in range(2, num_rows + 1, 2):
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
                    requests.append(self._create_request(row - 1, num_cols, sheet_id, formatting_config['background_color'], True, 0))  # col_index=0 for entire row

        if 'bold_rows' in formatting_config:
            for row in formatting_config['bold_rows']:
                requests.append(self._create_request(row - 1, num_cols, sheet_id, {'textFormat': {'bold': True}}, True, 0))  # col_index=0 for entire row
        return requests


    def _apply_conditional_formatting(self, conditional_formats, sheet_id, num_cols, values):
        requests = []
        header = values[0]
        for cond_format in conditional_formats:
            column_name = cond_format.get('column')
            if not column_name:
                print("Missing 'column' key in conditional formatting.")
                continue

            values_config = cond_format.get('values')
            condition_func = cond_format.get('condition')
            entire_row = cond_format.get('entire_row', False)
            format_style = cond_format.get('format')

            try:
                col_index = header.index(column_name)
            except ValueError:
                print(f"Column '{column_name}' not found for conditional formatting.")
                continue

            for i, row in enumerate(values[1:], 1):
                cell_value = row[col_index]
                try:
                    if values_config and cell_value in values_config:
                        requests.append(self._create_request(i, num_cols, sheet_id, values_config[cell_value], entire_row, col_index))
                    elif condition_func:
                        try:
                            int_cell_value = int(cell_value)  # Try converting to int
                            if condition_func(int_cell_value):
                                requests.append(self._create_request(i, num_cols, sheet_id, format_style, entire_row, col_index))
                        except ValueError:
                            print(f"Non-integer value found in '{column_name}' column: {cell_value}. Skipping conditional formatting for this row.")
                except TypeError as e:
                    print(f"Error applying condition: {e}. Skipping row {i + 1}")
        return requests

   
