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
        Conflicting properties are merged, and new properties are added.
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
                elif isinstance(merged_format[key], list) and isinstance(value, list):
                    # Merge lists (e.g., format for case-specific formatting)
                    merged_format[key].extend(value)
                else:
                    # Overwrite non-dict and non-list properties
                    merged_format[key] = value
        return merged_format

    def _update_cache(self, row_index: int, col_index: int, new_format: Dict[str, Any]):
        """
        Update the formatting cache for a specific cell.
        """
        current_format = self.formatting_cache[row_index][col_index]
        merged_format = self._merge_formatting(current_format, new_format)
        self.formatting_cache[row_index][col_index] = merged_format

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
                                    for extra_col in cond_format['extra_columns']:
                                        extra_col_index = header.index(extra_col)
                                        self._update_cache(i, extra_col_index, format_style)
                                if i == 1:  # Only print debug output for the first instance
                                    print(f"Applied conditional formatting to cell ({i}, {col_index+1}): {format_style}")
                            else:
                                if i == 1:  # Only print debug output for the first instance
                                    print(f"Skipped conditional formatting for cell ({i}, {col_index+1})")
