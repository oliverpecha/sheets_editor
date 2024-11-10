from typing import Any, Dict, List

class SheetFormatter:
    def __init__(self):
      pass #Needed if no instance attributes
        
    def _create_request(self, row_index, num_cols, sheet_id, format_style, entire_row, col_index):
        """
        Creates a formatting request for a specific cell or the entire row.
        """
        start_col = 0 if entire_row else col_index
        end_col = num_cols if entire_row else col_index + 1
    
        user_entered_format = {}
    
        # Handle background color and other format styles
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
        # Debug print to check what conditional formats are being passed
        print(f"Conditional Formats: {conditional_formats}")

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
        """Applies conditional formatting based on the provided conditions."""
        requests = []
        header = values[0]  # Assuming the first row contains headers (column names)
    
        for cond_format in conditional_formats:
            conditions = cond_format.get('conditions', [])
            format_name = cond_format.get('name', 'Unnamed Format')
            formatting_type = cond_format.get('type', 'case_specific')  # Default to case_specific
            print(f"Processing conditional format '{format_name}' of type '{formatting_type}' with conditions: {conditions}")
    
            if not conditions:
                print("No conditions provided for conditional formatting.")
                continue
    
            # Iterate over the data rows (skipping the header row)
            for i, row in enumerate(values[1:], 1):  # Start from the second row (data rows)
                if formatting_type == 'case_specific':
                    for index, condition in enumerate(conditions):
                        column_name = condition.get('column')
                        condition_func = condition.get('condition')
                        format_style = cond_format.get('format')[index]  # Get format for the current condition
    
                        if column_name not in header:
                            print(f"Column '{column_name}' not found in the header.")
                            continue
    
                        col_index = header.index(column_name)  # Find the column index
                        cell_value = row[col_index]  # Get the cell's value
    
                        if condition_func(cell_value):
                            print(f"Applying case-specific formatting for '{format_name}' on row {i + 1}: {cell_value}")
                            requests.append(self._create_request(i, num_cols, sheet_id, format_style, False, col_index))
    
                elif formatting_type == 'all_conditions':
                    all_conditions_met = True
                    for condition in conditions:
                        column_name = condition.get('column')
                        condition_func = condition.get('condition')
    
                        if column_name not in header:
                            print(f"Column '{column_name}' not found in the header.")
                            all_conditions_met = False
                            break
    
                        col_index = header.index(column_name)  # Find the column index
                        cell_value = row[col_index]  # Get the cell's value
    
                        if not condition_func(cell_value):
                            all_conditions_met = False
                            break
    
                    if all_conditions_met:
                        print(f"Applying all-conditions formatting for '{format_name}' to row {i + 1}")
                        requests.append(self._create_request(i, num_cols, sheet_id, cond_format.get('format'), True, 0))
    
        return requests

   
