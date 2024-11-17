from typing import Any, Dict

class SheetFormatter:
    def __init__(self):
        """Initialize the formatting cache."""
        self.formatting_cache = {}
        self.debug_enabled = True  # Control debug output globally

    def format_worksheet(self, worksheet, formatting_config=None, conditional_formats=None):
        """
        Main public method to format a worksheet with both absolute and conditional formatting.
        
        Args:
            worksheet: The worksheet object to format
            formatting_config: Dictionary containing absolute formatting rules
            conditional_formats: List of conditional formatting rules
        """
        # Get all values from worksheet
        values = worksheet.get_all_values()
        if not values:
            print("No values found in worksheet")
            return []

        # Get worksheet properties
        num_rows = len(values)
        num_cols = len(values[0]) if values else 0
        sheet_id = worksheet.id

        # Initialize the formatting cache
        self._initialize_cache(num_rows, num_cols)

        # Apply absolute formatting if configured
        if formatting_config:
            self._apply_absolute_formatting(formatting_config, sheet_id, num_rows, num_cols)

        # Apply conditional formatting if configured
        if conditional_formats:
            self._apply_conditional_formatting(conditional_formats, sheet_id, values)

        # Generate final formatting requests
        return self._generate_requests_from_cache(sheet_id, num_cols)

    def _initialize_cache(self, num_rows: int, num_cols: int):
        """Initialize the formatting cache for all cells in the sheet."""
        self.formatting_cache = {
            row: {col: {} for col in range(num_cols)} for row in range(num_rows)
        }
        if self.debug_enabled:
            print(f"\nInitialized formatting cache for {num_rows} rows and {num_cols} columns")

    def _merge_formatting(self, current_format: Dict[str, Any], new_format: Dict[str, Any]) -> Dict[str, Any]:
        """
        Merge new formatting into the current formatting for a cell.
        Preserves existing formatting while adding new properties.
        """
        merged_format = current_format.copy()
        
        for key, value in new_format.items():
            if isinstance(value, dict) and key in merged_format:
                # Deep merge for nested properties like backgroundColor or textFormat
                merged_format[key] = {**merged_format[key], **value}
            else:
                # Direct assignment for non-dict properties
                merged_format[key] = value
        
        return merged_format

    def _update_cache(self, row_index: int, col_index: int, new_format: Dict[str, Any]):
        """Update the formatting cache for a specific cell."""
        if row_index not in self.formatting_cache:
            self.formatting_cache[row_index] = {}
        if col_index not in self.formatting_cache[row_index]:
            self.formatting_cache[row_index][col_index] = {}
            
        current_format = self.formatting_cache[row_index][col_index]
        self.formatting_cache[row_index][col_index] = self._merge_formatting(current_format, new_format)

    def _generate_requests_from_cache(self, sheet_id: int, num_cols: int):
        """Generate batch update requests from the formatting cache."""
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
                            entire_row=False,
                            col_index=col_index
                        )
                    )
        return requests

    def _create_request(self, row_index: int, num_cols: int, sheet_id: int, 
                       format_style: Dict[str, Any], entire_row: bool, col_index: int):
        """Create a formatting request for the Google Sheets API."""
        start_col = 0 if entire_row else col_index
        end_col = num_cols if entire_row else col_index + 1

        # Properly structure the format request
        user_entered_format = {}
        
        # Handle backgroundColor separately
        if isinstance(format_style.get('backgroundColor'), dict):
            user_entered_format['backgroundColor'] = format_style['backgroundColor']
        
        # Handle textFormat
        if 'textFormat' in format_style:
            user_entered_format['textFormat'] = format_style['textFormat']
            
        # Add any other formatting properties
        for key, value in format_style.items():
            if key not in ['backgroundColor', 'textFormat']:
                user_entered_format[key] = value

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
                "fields": "userEnteredFormat"
            }
        }

    def _apply_absolute_formatting(self, formatting_config: Dict[str, Any], sheet_id: int, 
                                 num_rows: int, num_cols: int):
        """Apply absolute formatting configurations."""
        if self.debug_enabled:
            print("\nApplying absolute formatting...")
            
        if formatting_config.get('alternate_rows'):
            for row in range(1, num_rows, 2):
                for col in range(num_cols):
                    self._update_cache(row, col, {'backgroundColor': formatting_config['background_color']})
                    
        if 'bold_rows' in formatting_config:
            for row in formatting_config['bold_rows']:
                for col in range(num_cols):
                    self._update_cache(row - 1, col, {'textFormat': {'bold': True}})

    def _apply_conditional_formatting(self, conditional_formats: list, sheet_id: int, values: list):
        """Apply conditional formatting based on the provided rules."""
        header = values[0]
        num_cols = len(header)
        debug_row = 1  # Track first data row for debugging

        if self.debug_enabled:
            print("\n=== Conditional Formatting Debug ===")
            print(f"First data row: {values[1]}")

        for format_rule in conditional_formats:
            format_type = format_rule.get('type', 'all_conditions')
            
            if self.debug_enabled:
                print(f"\nProcessing format rule: {format_rule.get('name', 'Unnamed')}")
                print(f"Format type: {format_type}")

            if format_type == 'case_specific':
                self._apply_case_specific_formatting(format_rule, header, values, debug_row)
            else:  # all_conditions
                self._apply_all_conditions_formatting(format_rule, header, values, num_cols, debug_row)

    def _apply_case_specific_formatting(self, format_rule, header, values, debug_row):
        """Handle case-specific conditional formatting."""
        for i, (condition, format_style) in enumerate(zip(format_rule['conditions'], format_rule['format'])):
            column_name = condition['column']
            if column_name in header:
                col_index = header.index(column_name)
                condition_func = condition['condition']

                for row_idx, row in enumerate(values[1:], 1):
                    cell_value = row[col_index]
                    if condition_func(cell_value):
                        current_format = self.formatting_cache[row_idx][col_index].copy()
                        merged_format =
