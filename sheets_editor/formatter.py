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

        values = worksheet.get_all_values()
        if not values:
            return

        num_rows = len(values)
        num_cols = len(values[0])

        # Initialize the formatting cache
        self._initialize_cache(num_rows, num_cols)

        # Apply absolute formatting
        if formatting_config:
            self._apply_absolute_formatting(formatting_config, num_rows, num_cols)

        # Apply conditional formatting
        if conditional_formats:
            self._apply_conditional_formatting(conditional_formats, values)

        # Generate batch requests from the cache and update the worksheet
        self._update_worksheet_from_cache(worksheet)

    def _apply_absolute_formatting(self, formatting_config, num_rows, num_cols):
        """Apply absolute formatting (e.g., alternate rows, bold rows)."""
        if formatting_config.get('alternate_rows'):
            for row in range(1, num_rows, 2):  # Apply to even rows (1-based indexing)
                for col in range(num_cols):
                    self._apply_format_to_cache(row, col, formatting_config['background_color'])

        if 'bold_rows' in formatting_config:
            for row in formatting_config['bold_rows']:
                for col in range(num_cols):
                    self._apply_format_to_cache(row - 1, col, {'textFormat': {'bold': True}})

    def _apply_conditional_formatting(self, conditional_formats, values):
        """Apply conditional formatting based on the provided rules."""
        header = values[0]
        for cond_format in conditional_formats:
            for i, row in enumerate(values[1:], 1):  # Skip header row
                conditions_met = self._check_conditions(cond_format['conditions'], row, header)
                if conditions_met:
                    for col_index, cell_value in enumerate(row):
                        if cond_format.get('entire_row', False) or header[col_index] in [c['column'] for c in cond_format['conditions']]:
                            self._apply_format_to_cache(i, col_index, cond_format['format'])

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
                "fields": "userEnteredFormat(backgroundColor,textFormat)"  # Specify fields to update
            }
        }
