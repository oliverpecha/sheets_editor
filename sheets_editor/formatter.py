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
        Conflicting properties will be overwritten by the new format.
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

    def _update_cache(self, row_index: int, col_index: int, new_format: Dict[str, Any], apply_only_if_empty: bool = False):
        """
        Update the formatting cache for a specific cell.
        If `apply_only_if_empty` is True, new formatting will only be applied to empty formatting fields.
        """
        current_format = self.formatting_cache[row_index][col_index]

        # Debug: Print current and new formatting
        print(f"Updating cache: row={row_index}, col={col_index}")
        print(f"Current format: {current_format}")
        print(f"New format: {new_format}")

        # Merge the new format with the current format
        for key, value in new_format.items():
            if key not in current_format:
                # If the key doesn't exist, add it
                current_format[key] = value
            else:
                if isinstance(current_format[key], dict) and isinstance(value, dict):
                    # Merge dictionaries (e.g., backgroundColor, textFormat)
                    for sub_key, sub_value in value.items():
                        if sub_key not in current_format[key] or not apply_only_if_empty:
                            print(f"Updating sub-key: {sub_key} with value: {sub_value}")
                            current_format[key][sub_key] = sub_value
                elif not apply_only_if_empty:
                    # Overwrite non-dict properties if `apply_only_if_empty` is False
                    print(f"Overwriting key: {key} with value: {value}")
                    current_format[key] = value

        # Debug: Print updated formatting
        print(f"Updated format: {self.formatting_cache[row_index][col_index]}")

    def _generate_requests_from_cache(self, sheet_id: int, num_cols: int):
        """
        Generate batch update requests from the formatting cache.
        This method converts the cached formatting for each cell into Google Sheets API requests.
        """
        requests = []
        for row_index, row in self.formatting_cache.items():
            for col_index, cell_format in row.items():
                if cell_format:  # Only generate requests for cells with formatting
                    print(f"Generating request for row={row_index}, col={col_index}, format={cell_format}")
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
            print(f"\nApplying conditional format: {cond_format['name']}")
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
                                print(f"Condition met for row={i}, col={col_index}, value={cell_value}")
                                self._update_cache(i, col_index, format_style)
            elif formatting_type == 'all_conditions':
                # Handle all-conditions formatting
                for i, row in enumerate(values[1:], 1):  # Skip header row
                    conditions_met = self._check_conditions(cond_format['conditions'], row, header)
                    if conditions_met:
                        print(f"All conditions met for row={i}")
                        if cond_format.get('entire_row', False):
                            for col_index in range(num_cols):
                                self._update_cache(i, col_index, cond_format['format'], apply_only_if_empty=False)
                        else:
                            for condition in cond_format['conditions']:
                                column_name = condition['column']
                                if column_name in header:
                                    col_index = header.index(column_name)
                                    self._update_cache(i, col_index, cond_format['format'])

            # Debug: Print cache state after applying the rule
            print(f"Cache state after applying {cond_format['name']}:")
            import json
            print(json.dumps(self.formatting_cache, indent=2))

    def _check_conditions(self, conditions: list, row: list, header: list) -> bool:
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

        # Apply absolute formatting
        if formatting_config:
            self._apply_absolute_formatting(formatting_config, sheet_id, num_rows, num_cols)

        # Apply conditional formatting
        if conditional_formats:
            self._apply_conditional_formatting(conditional_formats, sheet_id, values)

        # Generate batch requests from the cache
        requests = self._generate_requests_from_cache(sheet_id, num_cols)

        # Debug: Print all final requests
        print("\nFinal batch requests:")
        import json
        print(json.dumps(request
