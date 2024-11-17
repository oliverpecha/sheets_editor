from typing import Any, Dict

class SheetFormatter:
    def __init__(self):
        self.formatting_cache = {}  # Initialize the formatting cache

    def _initialize_cache(self, num_rows: int, num_cols: int):
        """Initialize the formatting cache for all cells in the sheet."""
        self.formatting_cache = {
            row: {col: {} for col in range(num_cols)} for row in range(num_rows)
        }

    def _merge_formatting(self, current_format: Dict[str, Any], new_format: Dict[str, Any]) -> Dict[str, Any]:
        """
        Merge new formatting into the current formatting for a cell.
        Conflicting properties are overwritten by the new format.
        """
        merged_format = current_format.copy()
        for key, value in new_format.items():
            if key not in merged_format:
                # Add new property if it doesn't exist
                merged_format[key] = value
            else:
                if isinstance(merged_format[key], dict) and isinstance(value, dict):
                    # Merge dictionaries (e.g., backgroundColor, textFormat)
                    merged_format[key].update(value)
                else:
                    # Overwrite non-dict properties
                    merged_format[key] = value
        return merged_format

    def _update_cache(self, row_index: int, col_index: int, new_format: Dict[str, Any]):
        """
        Update the formatting cache for a specific cell.
        """
        current_format = self.formatting_cache[row_index][col_index]
        self.formatting_cache[row_index][col_index] = self._merge_formatting(current_format, new_format)

        # Debugging console output for the first few rows
        if row_index < 5:  # Output formatting for the first 5 rows
            print(f"Formatting for row {row_index}:")
            for col, cell_format in self.formatting_cache[row_index].items():
                print(f"  Column {col}: {cell_format}")

    def _generate_requests_from_cache(self, sheet_id: int, num_cols: int):
        """
        Generate batch update requests from the formatting cache.
        This method converts the cached formatting for each cell into Google Sheets API requests.
        """
        requests = []
        for row_index, row in self.formatting_cache.items():
            for col_index, cell_format in row.items():
                if cell_format:  # Only generate requests for cells with formatting
                    requests.append(
                        self._create_request(
                            row_index=row_index,
                            num_cols=num_cols,
                            sheet_id=sheet_id,
                            format_style=cell_format,
                            entire_row=False,  # Specific cell formatting
                            col_index=col_index
                        )
                    )
        return requests

    def _create_request(self, row_index: int, num_cols: int, sheet_id: int, format_style: Dict[str, Any], entire_row: bool, col_index: int):
        """
        Creates a formatting request for a specific cell or the entire row.
        """
        start_col = 0 if entire_row else col_index
        end_col = num_cols if entire_row else col_index + 1

        user_entered_format = {}

        # Handle backgroundColor and other format styles
        if "red" in format_style and "green" in format_style and "blue" in format_style:
            user_entered_format["backgroundColor"] = format_style  # Properly nest the color object
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

    def _apply_absolute_formatting(self, formatting_config: Dict[str, Any], sheet_id: int, num_rows: int, num_cols: int):
        """Apply absolute formatting (e.g., alternate rows, bold rows)."""
        if formatting_config.get('alternate_rows'):
            for row in range(1, num_rows, 2):  # Apply to every other row (even rows)
                for col in range(num_cols):
                    self._update_cache(row, col, formatting_config['background_color'])

        if 'bold_rows' in formatting_config:
            for row in formatting_config['bold_rows']:
                for col in range(num_cols):
                    self._update_cache(row - 1, col, {'textFormat': {'bold': True}})

    def _apply_conditional_formatting(self, conditional_formats: list, sheet_id: int, values: list):
        """Apply conditional formatting based on the provided rules."""
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
                        for i, row in enumerate(values[1:], 1):  # Skip header row
                            cell_value = row[col_index]
                            if condition_func(cell_value):  # Check if the condition is met
                                self._update_cache(i, col_index, format_style)
                                if cond_format.get('entire_row', False):
                                    for extra_col in range(num_cols):
                                        if extra_col != col_index:
                                            self._update_cache(i, extra_col, format_style)
                                if 'extra_columns' in cond_format:
                                    extra_columns_indices = [header.index(col) for col in cond_format['extra_columns']]
                                    for extra_col in extra_columns_indices:
                                        self._update_cache(i, extra_col, format_style)

            elif formatting_type == 'all_conditions':
                # Find rows where all conditions are met
                matching_rows = []
                for i, row in enumerate(values[1:], 1):  # Skip header row
                    if all(condition_func(row[header.index(column_name)]) for condition, condition_func in zip(cond_format['conditions'], cond_format['condition_funcs'])):
                        matching_rows.append(i)

                # Merge formatting for matching rows
                for row in matching_rows:
                    for col in range(num_cols):
                        if cond_format.get('entire_row', False) or col in extra_columns_indices:
                            self._update_cache(row, col, cond_format['format'])
                        if 'extra_columns' in cond_format:
                            extra_columns_indices = [header.index(col) for col in cond_format['extra_columns']]
                            for extra_col in extra_columns_indices:
                                self._update_cache(row, extra_col, cond_format['format'])


    def format_worksheet(self, worksheet, values):
        """
        Apply formatting to a worksheet based on the cached formatting rules.
    
        :param worksheet: The gspread Worksheet object to format.
        :param values: A 2D list representing the sheet data (including headers).
        """
        sheet_id = worksheet.id
        num_rows = len(values)
        num_cols = len(values[0]) if values else 0
    
        # Initialize the formatting cache
        self._initialize_cache(num_rows, num_cols)
    
        # Apply absolute formatting (if any)
        if self.absolute_formatting:
            self._apply_absolute_formatting(self.absolute_formatting, sheet_id, num_rows, num_cols)
    
        # Apply conditional formatting
        if self.conditional_formats:
            self._apply_conditional_formatting(self.conditional_formats, sheet_id, values)
    
        # Generate and apply the formatting requests
        requests = self._generate_requests_from_cache(sheet_id, num_cols)
        worksheet.batch_update(requests)
