from typing import Any, Dict, List

class SheetFormatter:
    def __init__(self):
        self.formatting_cache = {}  # Initialize empty formatting cache

    def _initialize_cache(self, num_rows, num_cols):
        """Initialize the formatting cache for the entire sheet."""
        self.formatting_cache = {
            row: {col: {} for col in range(num_cols)} for row in range(num_rows)
        }

    def _merge_formatting(self, current_format, new_format):
        """
        Merge new formatting into the current formatting for a cell.
        Conflicting properties will be overwritten by the new format.
        """
        merged_format = current_format.copy()
        for key, value in new_format.items():
            if key not in merged_format or not isinstance(merged_format[key], dict):
                # Add new property or overwrite non-dict properties
                merged_format[key] = value
            else:
                # Merge dictionaries (e.g., backgroundColor or textFormat)
                merged_format[key].update(value)
        return merged_format

    def _apply_format_to_cache(self, row_index, col_index, new_format):
        """Apply a new format to a specific cell in the cache."""
        current_format = self.formatting_cache[row_index][col_index]
        self.formatting_cache[row_index][col_index] = self._merge_formatting(current_format, new_format)

    def format_worksheet(self, worksheet, formatting_config=None, conditional_formats=None):
        """Apply absolute and conditional formatting to the worksheet."""
        if not formatting_config and not conditional_formats:
            return
    
        # Get all values to determine the number of rows and columns
        values = worksheet.get_all_values()
        if not values:
            return
    
        num_rows = len(values)
        num_cols = len(values[0]) if values else 0
        sheet_id = worksheet._properties['sheetId']  # Get the sheet ID
    
        # Initialize cache for the entire sheet
        self._initialize_cache(num_rows, num_cols)
    
        # Initialize the batch requests list
        requests = []
    
        # Apply absolute formatting
        if formatting_config:
            self._apply_absolute_formatting(formatting_config, sheet_id, num_rows, num_cols)
    
        # Apply conditional formatting
        if conditional_formats:
            self._apply_conditional_formatting(conditional_formats, sheet_id, values)
    
        # Generate batch requests from the cache
        requests = self._generate_requests_from_cache(sheet_id, num_cols)
    
        # Send batch update to Google Sheets API
        if requests:
            try:
                worksheet.spreadsheet.batch_update({"requests": requests})
            except Exception as e:
                print(f"Error applying formatting: {e}")
                raise

    def _apply_absolute_formatting(self, formatting_config, sheet_id, num_rows, num_cols):
        """Apply absolute formatting (e.g., alternate rows, bold rows)."""
        requests = []
        if formatting_config.get('alternate_rows'):
            # Apply to every other row (even rows, 1-based indexing)
            for row in range(1, num_rows, 2):  # Skip every other row
                requests.append(
                    self._create_request(
                        row_index=row - 1,  # Convert to 0-based indexing
                        num_cols=num_cols,
                        sheet_id=sheet_id,
                        format_style=formatting_config['background_color'],
                        entire_row=True,  # Apply formatting to the entire row
                        col_index=0  # This value is ignored because `entire_row=True`
                    )
                )
    
        if 'bold_rows' in formatting_config:
            # Apply bold formatting to specific rows
            for row in formatting_config['bold_rows']:
                requests.append(
                    self._create_request(
                        row_index=row - 1,  # Convert to 0-based indexing
                        num_cols=num_cols,
                        sheet_id=sheet_id,
                        format_style={'textFormat': {'bold': True}},
                        entire_row=True,  # Apply formatting to the entire row
                        col_index=0  # This value is ignored because `entire_row=True`
                    )
                )
        return requests

    def _apply_conditional_formatting(self, conditional_formats, sheet_id, values):
        """Apply conditional formatting based on the provided rules."""
        requests = []
        header = values[0]  # Assuming the first row is the header
        num_cols = len(header)
    
        for cond_format in conditional_formats:
            formatting_type = cond_format.get('type', 'all_conditions')  # Default to 'all_conditions'
    
            if formatting_type == 'case_specific':
                # Handle case-specific formatting
                for condition, format_style in zip(cond_format['conditions'], cond_format['format']):
                    column_name = condition['column']
                    condition_func = condition['condition']
                    if column_name in header:
                        col_index = header.index(column_name)
                        for i, row in enumerate(values[1:], 1):  # Skip header row, start from the second row
                            cell_value = row[col_index]
                            if condition_func(cell_value):  # Check if the condition is met
                                requests.append(
                                    self._create_request(
                                        row_index=i,
                                        num_cols=num_cols,
                                        sheet_id=sheet_id,
                                        format_style=format_style,
                                        entire_row=False,  # Apply to the specific column
                                        col_index=col_index
                                    )
                                )
            elif formatting_type == 'all_conditions':
                # Handle all-conditions formatting
                for i, row in enumerate(values[1:], 1):  # Skip header row, start from the second row
                    conditions_met = self._check_conditions(cond_format['conditions'], row, header)
                    if conditions_met:
                        # Apply formatting to the entire row if specified
                        if cond_format.get('entire_row', False):
                            requests.append(
                                self._create_request(
                                    row_index=i,
                                    num_cols=num_cols,
                                    sheet_id=sheet_id,
                                    format_style=cond_format['format'],
                                    entire_row=True,  # Apply to the entire row
                                    col_index=0  # Ignored when entire_row=True
                                )
                            )
                        else:
                            # Apply formatting only to the columns specified in the conditions
                            for condition in cond_format['conditions']:
                                column_name = condition['column']
                                if column_name in header:
                                    col_index = header.index(column_name)
                                    requests.append(
                                        self._create_request(
                                            row_index=i,
                                            num_cols=num_cols,
                                            sheet_id=sheet_id,
                                            format_style=cond_format['format'],
                                            entire_row=False,  # Apply to a specific column
                                            col_index=col_index
                                        )
                                    )
                        # Apply formatting to extra columns if specified
                        if 'extra_columns' in cond_format:
                            for extra_column in cond_format['extra_columns']:
                                if extra_column in header:
                                    extra_col_index = header.index(extra_column)
                                    requests.append(
                                        self._create_request(
                                            row_index=i,
                                            num_cols=num_cols,
                                            sheet_id=sheet_id,
                                            format_style=cond_format['format'],
                                            entire_row=False,  # Apply to the specific extra column
                                            col_index=extra_col_index
                                        )
                                    )
        return requests
        
    def _check_conditions(self, conditions, row, header):
        """Check if all conditions are met for a given row."""
        for condition in conditions:
            column_name = condition['column']
            condition_func = condition['condition']
            if column_name not in header:
                return False
            col_index = header.index(column_name)
            if not condition_func(row[col_index]):
                return False
        return True

    def _update_worksheet_from_cache(self, worksheet):
        """Generate batch requests from the formatting cache and update the worksheet."""
        requests = []
        for row_index, row in self.formatting_cache.items():
            for col_index, cell_format in row.items():
                if cell_format:  # Only create requests for cells with formatting
                    requests.append(self._create_request(row_index, col_index, worksheet._properties['sheetId'], cell_format))

        if requests:
            # Send batch update to Google Sheets API
            worksheet.spreadsheet.batch_update({"requests": requests})

    def _create_request(self, row_index, num_cols, sheet_id, format_style, entire_row, col_index):
        """
        Creates a formatting request for a specific cell or the entire row.
        """
        start_col = 0 if entire_row else col_index
        end_col = num_cols if entire_row else col_index + 1
    
        user_entered_format = {}
    
        # Handle backgroundColor and other format styles
        if "red" in format_style and "green" in format_style and "blue" in format_style:
            user_entered_format["backgroundColor"] = format_style  # Nest the color object properly
        else:
            user_entered_format.update(format_style)  # Add other formatting styles
    
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
                "fields": "userEnteredFormat"  # Specify all fields to update
            }
        }
