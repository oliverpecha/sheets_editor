from typing import Any, Dict, List
#temp
class SheetFormatter:
    def __init__(self):
        pass  # Needed if no instance attributes

    def _create_request(self, row_index, num_cols, sheet_id, format_style, entire_row, col_index, existing_format=None):
        """Creates a formatting request, correctly handling colors."""
        start_col = 0 if entire_row else col_index
        end_col = num_cols if entire_row else col_index + 1
    
        user_entered_format = existing_format.copy() if existing_format else {}
    
        if isinstance(format_style, dict):
            # Correctly format colors:
            if all(k in format_style for k in ("red", "green", "blue")):
                color = format_style.copy()  # Don't modify the original format_style
                user_entered_format['backgroundColor'] = {'red': color.pop('red', 0), 'green': color.pop('green', 0), 'blue': color.pop('blue', 0), 'alpha': color.pop('alpha', 1)} #added alpha
                user_entered_format.update(color) #add remaining style key-values
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

        print(f"Conditional Formats: {conditional_formats}")

        try:
            values = worksheet.get_all_values()
            if not values:
                return

            header = values[0]
            num_cols = len(header)
            requests = []

            if formatting_config:
                requests.extend(self._apply_absolute_formatting(formatting_config, worksheet._properties['sheetId'], len(values), num_cols, values))

            if conditional_formats:
                requests.extend(self._apply_conditional_formatting(conditional_formats, worksheet._properties['sheetId'], values, worksheet))

            if requests:
                print(json.dumps(requests, indent=4))  # Print the requests for debugging
                worksheet.spreadsheet.batch_update({"requests": requests})

        except Exception as e:
            print(f"Error in formatting: {e}")
            raise



    def _apply_absolute_formatting(self, formatting_config, sheet_id, num_rows, num_cols, values):
        requests = []

        # Check data types of colors *before* building requests:
        if formatting_config.get('background_color'):  # Check if background_color is present
            for color_key in ('red', 'green', 'blue'):
                if color_key in formatting_config['background_color']:
                    color_value = formatting_config['background_color'][color_key]
                    print(f"Value type for {color_key}: {type(color_value)}")  # Check and print type
                    if not isinstance(color_value, (int, float)):
                        formatting_config['background_color'][color_key] = float(color_value) # Attempt to convert to float if necessary

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

    def _apply_conditional_formatting(self, conditional_formats, sheet_id, values, worksheet):
        """Applies conditional formatting, with added type checking for colors."""
        requests = []
        header = values[0]
        num_cols = len(header)

        for cond_format in conditional_formats:
            conditions = cond_format.get('conditions', [])
            format_name = cond_format.get('name', 'Unnamed Format')
            formatting_type = cond_format.get('type', 'case_specific')
            entire_row = cond_format.get('entire_row', False)
            extra_columns = cond_format.get('extra_columns', [])

            print(f"Processing conditional format '{format_name}' of type '{formatting_type}' with conditions: {conditions}")

            if not conditions:
                print("No conditions provided for conditional formatting.")
                continue

            format_style = cond_format.get('format')

            # Check and convert color data types before creating requests:
            if isinstance(format_style, dict) and 'backgroundColor' in format_style:
                for color_key in ('red', 'green', 'blue'):
                    if color_key in format_style['backgroundColor']:
                        color_value = format_style['backgroundColor'][color_key]
                        print(f"Value type for {color_key}: {type(color_value)}")
                        if not isinstance(color_value, (int, float)):
                            try:
                                format_style['backgroundColor'][color_key] = float(color_value)
                            except (ValueError, TypeError):
                                print(f"Warning: Could not convert {color_key} value to float: {color_value}")

            if isinstance(format_style, list):  # Handle list of formats for case-specific
                for fmt in format_style:
                    if isinstance(fmt, dict) and 'backgroundColor' in fmt:
                        for color_key in ('red', 'green', 'blue'):
                            if color_key in fmt['backgroundColor']:
                                color_value = fmt['backgroundColor'][color_key]
                                print(f"Value type for {color_key}: {type(color_value)}")
                                if not isinstance(color_value, (int, float)):
                                    try:
                                        fmt['backgroundColor'][color_key] = float(color_value)
                                    except (ValueError, TypeError):
                                        print(f"Warning: Could not convert {color_key} value to float: {color_value}")

            for i, row in enumerate(values[1:], 1):
                try:
                    existing_row_formats = worksheet.get_row_formats(i + 1)
        
                    if existing_row_formats:
                        first_cell_format = existing_row_formats[0].get("userEnteredFormat") # Access userEnteredFormat directly
                        existing_format = first_cell_format.copy() if first_cell_format else None
                    else:
                        existing_format = None
        
                except Exception as e: #Handle exceptions during format retrieval
                    print(f"Error getting row formats: {e}")
                    existing_format = None # Default to no existing format if there's an error

                
                if formatting_type == 'case_specific':
                    self._apply_case_specific_formatting(requests, i, row, conditions, cond_format, header, sheet_id, existing_format)
                elif formatting_type == 'all_conditions':
                    self._apply_all_conditions_formatting(requests, i, row, conditions, cond_format, header, sheet_id, existing_format)

        return requests
       
    def _apply_case_specific_formatting(self, requests, row_index, row, conditions, cond_format, header, sheet_id, existing_format):
        """Applies case-specific formatting if conditions are met."""
        for index, condition in enumerate(conditions):
            column_name = condition.get('column')
            condition_func = condition.get('condition')
            format_style = cond_format.get('format')[index] if isinstance(cond_format.get('format'), list) else cond_format.get('format')
    
            # Ensure the column exists
            if column_name not in header:
                print(f"Column '{column_name}' not found in the header.")
                continue
    
            col_index = header.index(column_name)  # Find the column index
            cell_value = row[col_index]  # Get the cell's value
    
            # Check if the condition is met
            if condition_func(cell_value):
                print(f"Applying case-specific formatting for '{cond_format['name']}' on row {row_index + 1}: {cell_value}")
                if cond_format.get('entire_row', False):
                    requests.append(self._create_request(row_index, len(header), sheet_id, format_style, True, 0, existing_format))  # Apply to entire row
                else:
                    requests.append(self._create_request(row_index, len(header), sheet_id, format_style, False, col_index, existing_format))  # Apply to specific column
    
    def _apply_all_conditions_formatting(self, requests, row_index, row, conditions, cond_format, header, sheet_id, existing_format):
        """Applies formatting if all conditions are met."""
        all_conditions_met = True
        for condition in conditions:
            column_name = condition.get('column')
            condition_func = condition.get('condition')
    
            # Ensure the column exists
            if column_name not in header:
                print(f"Column '{column_name}' not found in the header.")
                all_conditions_met = False
                break  # Exit the loop since one condition is not met
    
            col_index = header.index(column_name)  # Find the column index
            cell_value = row[col_index]  # Get the cell's value
    
            # Check if the condition is met
            if not condition_func(cell_value):
                all_conditions_met = False
                break
    
        # If all conditions are met, apply the formatting
        if all_conditions_met:
            print(f"Applying all-conditions formatting for '{cond_format['name']}' to row {row_index + 1}")
            format_style = cond_format.get('format')
            if cond_format.get('entire_row', False):
                requests.append(self._create_request(row_index, len(header), sheet_id, format_style, True, 0, existing_format.copy()))  # Apply to entire row
            else:
                # Apply to extra columns if specified, merging with existing formats
                for extra_col in cond_format.get('extra_columns', []):
                    if extra_col in header:
                        extra_col_index = header.index(extra_col)
                        requests.append(self._create_request(row_index, len(header), sheet_id, format_style, False, extra_col_index, existing_format.copy()))  # Apply to extra columns
