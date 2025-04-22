import gspread 
from typing import Any, Dict, List, Optional, Tuple 

#Version: "publish to production"

class SheetFormatter:
     def __init__(self, debug_enabled: bool = False): # Add debug_enabled parameter with a default
        """
        Initialize the SheetFormatter.

        Args:
            debug_enabled (bool, optional): If True, enables debug print statements.
                                            Defaults to False.
        """
        self.formatting_cache: Dict[int, Dict[int, Dict[str, Any]]] = {}
        self.debug_enabled = debug_enabled # Set the instance attribute from the parameter
        if self.debug_enabled:
            print("SheetFormatter initialized with DEBUG ENABLED.") # Optional: confirm debug mode

        
    def export_table(self, data, version, sheet_name, spreadsheet, formatting=None, conditional_formats=None):
        """
        Export data to a Google Sheets worksheet with formatting support.
        
        Args:
            data: List of dictionaries containing the data to export
            version: Version string for tracking
            sheet_name: Name of the sheet to create/update
            spreadsheet: Google Sheets spreadsheet object
            formatting: Optional dictionary containing absolute formatting rules
            conditional_formats: Optional list of conditional formatting rules
        """
        try:
            # Create or get worksheet
            try:
                worksheet = spreadsheet.worksheet(sheet_name)
                print(f"Found existing worksheet: {sheet_name}")
            except gspread.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="26")
                print(f"Created new worksheet: {sheet_name}")
    
            # Handle empty data case
            if not data:
                worksheet.clear()
                worksheet.update('A1', [['No data available']])
                print("No data to export")
                return worksheet
    
            # Convert data to DataFrame
            df = pd.DataFrame(data)
            
            # Apply column ignoring if specified in config
            if hasattr(self.config, 'ignore_columns') and self.config.ignore_columns:
                df = df.drop(columns=self.config.ignore_columns, errors='ignore')
                print(f"Ignored columns: {self.config.ignore_columns}")
    
            # Get headers and values
            headers = df.columns.tolist()
            values = df.values.tolist()
    
            # Prepare data for upload with headers
            all_values = [headers] + values
    
            # Clear existing content
            worksheet.clear()
    
            # Update the worksheet with data
            worksheet.update('A1', all_values, value_input_option='RAW')
            print(f"Updated worksheet with {len(values)} rows of data")
    
            # Apply formatting if provided
            if formatting or conditional_formats:
                print("Applying formatting...")
                try:
                    self.formatter.format_worksheet(
                        worksheet=worksheet,
                        formatting_config=formatting,
                        conditional_formats=conditional_formats
                    )
                except Exception as e:
                    print(f"Error applying formatting: {str(e)}")
                    raise
    
            # Auto-resize columns to fit content
            try:
                worksheet.columns_auto_resize(0, len(headers))
                print("Auto-resized columns")
            except Exception as e:
                print(f"Warning: Could not auto-resize columns: {str(e)}")
    
            # Return the worksheet for potential further operations
            return worksheet
    
        except Exception as e:
            print(f"Error in export_table: {str(e)}")
            raise

    # In your SheetFormatter class

    def format_worksheet(self,
                         worksheet: gspread.Worksheet,
                         # Renamed for clarity, this is for general rules
                         general_rules_config: Optional[Dict[str, Any]] = None,
                         # New parameter for targeted cell formats
                         targeted_cell_formats: Optional[Dict[Tuple[int, int], Dict[str, Any]]] = None,
                         conditional_formats: Optional[List[Dict[str, Any]]] = None):
        """
        Main public method to format a worksheet.
        Handles general rules, targeted cell formats, and conditional formatting.
        """
        if self.debug_enabled:
            print(f"\nSheetFormatter: Formatting worksheet: {worksheet.title} (ID: {worksheet.id})")
            print(f"SheetFormatter: Received general_rules_config: {general_rules_config}")
            print(f"SheetFormatter: Received targeted_cell_formats: {targeted_cell_formats}") # Log the new input
            print(f"SheetFormatter: Received conditional_formats count: {len(conditional_formats) if conditional_formats else 0}")

        values = worksheet.get_all_values()
        if not values:
            if self.debug_enabled:
                print("SheetFormatter: No values found in worksheet, skipping formatting.")
            return

        num_sheet_rows = len(values)
        num_sheet_cols = len(values[0]) if values else 0
        sheet_id = worksheet.id

        if num_sheet_rows == 0 or num_sheet_cols == 0:
            if self.debug_enabled:
                print("SheetFormatter: Worksheet is empty (0 rows or 0 cols), skipping formatting.")
            return

        try:
            self._initialize_cache(num_sheet_rows, num_sheet_cols) # Cache based on actual sheet size

            # Apply general absolute formatting rules first (e.g., alternate rows)
            if general_rules_config:
                self._apply_absolute_formatting(general_rules_config, sheet_id, num_sheet_rows, num_sheet_cols)

            # Apply targeted cell-specific formatting (can override general rules)
            if targeted_cell_formats: # Call the new method
                self._apply_targeted_formatting(targeted_cell_formats, sheet_id, num_sheet_rows, num_sheet_cols)

            # Apply conditional formatting
            if conditional_formats:
                self._apply_conditional_formatting(conditional_formats, sheet_id, values) # values is already all sheet values

            # Generate and apply requests from the populated cache
            requests = self._generate_requests_from_cache(sheet_id, num_sheet_cols)

            if self.debug_enabled:
                print(f"\nSheetFormatter: Generated {len(requests)} formatting requests from cache.")
                if requests:
                    print("SheetFormatter: Sample request:", requests[0]) # Keep sample logging

            if requests:
                try:
                    worksheet.spreadsheet.batch_update({"requests": requests})
                    if self.debug_enabled:
                        print("SheetFormatter: ✅ Successfully applied formatting requests via batch_update.")
                except gspread.exceptions.APIError as e:
                    print(f"SheetFormatter: ❌ APIError applying formatting requests: {e.response.text}")
                    raise
                except Exception as e:
                    print(f"SheetFormatter: ❌ Error applying formatting requests: {str(e)}")
                    raise
            elif self.debug_enabled:
                print("SheetFormatter: No formatting requests generated from cache to send.")

        except Exception as e:
            print(f"SheetFormatter: Error during formatting process: {str(e)}")
            raise
            
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
                        merged_format = self._merge_formatting(current_format, format_style)
                        self._update_cache(row_idx, col_index, merged_format)
    
                        if format_rule.get('entire_row', False):
                            # Apply to entire row
                            for col in range(len(header)):
                                if col != col_index:  # Skip the original column
                                    row_format = self.formatting_cache[row_idx][col].copy()
                                    merged_row_format = self._merge_formatting(row_format, format_style)
                                    self._update_cache(row_idx, col, merged_row_format)
    
                        if format_rule.get('extra_columns'):
                            # Apply to specified extra columns
                            for extra_col in format_rule['extra_columns']:
                                if extra_col in header:
                                    extra_col_idx = header.index(extra_col)
                                    extra_format = self.formatting_cache[row_idx][extra_col_idx].copy()
                                    merged_extra_format = self._merge_formatting(extra_format, format_style)
                                    self._update_cache(row_idx, extra_col_idx, merged_extra_format)
    
                        if self.debug_enabled and row_idx == debug_row:
                            print(f"\nCase-specific format applied:")
                            print(f"Column: {column_name}, Value: {cell_value}")
                            print(f"Format applied: {merged_format}")
    
    def _apply_targeted_formatting(self,
                                   targeted_formats: Dict[Tuple[int, int], Dict[str, Any]],
                                   sheet_id: int,
                                   num_sheet_rows: int, # Total rows in the sheet
                                   num_sheet_cols: int): # Total columns in the sheet
        """
        Applies formatting to specific cells based on (row_index, col_index) keys.
        Assumes row_index and col_index are 0-based relative to the data block (after header).
        """
        if self.debug_enabled:
            print(f"\nSheetFormatter: _apply_targeted_formatting received {len(targeted_formats)} rules.")

        if not targeted_formats:
            if self.debug_enabled:
                print("SheetFormatter: No targeted formats provided.")
            return

        # data_row_idx is 0-indexed *relative to the start of the data block* (i.e., after the header).
        # data_col_idx is 0-indexed *relative to the start of the exported columns*.
        header_offset = 1  # Because SheetsExporter adds one header row before data

        for (data_row_idx, data_col_idx), cell_style in targeted_formats.items():
            # Basic validation for the item structure
            if not (isinstance(data_row_idx, int) and isinstance(data_col_idx, int) and isinstance(cell_style, dict)):
                if self.debug_enabled:
                    print(f"SheetFormatter: Warning - Invalid item in targeted_formats: ({data_row_idx},{data_col_idx}): {cell_style}. Skipping.")
                continue

            # Convert data-relative indices to 0-based sheet indices
            actual_sheet_row_idx = data_row_idx + header_offset
            actual_sheet_col_idx = data_col_idx # Assuming data_col_idx is already the 0-based sheet column index

            # Boundary check against the actual sheet dimensions
            if not (0 <= actual_sheet_row_idx < num_sheet_rows and 0 <= actual_sheet_col_idx < num_sheet_cols):
                if self.debug_enabled:
                    print(f"SheetFormatter: Warning - Skipping out-of-bounds targeted format. "
                          f"Data coord: ({data_row_idx},{data_col_idx}) -> "
                          f"Sheet coord: ({actual_sheet_row_idx},{actual_sheet_col_idx}). "
                          f"Sheet dims: ({num_sheet_rows} R x {num_sheet_cols} C)")
                continue

            if self.debug_enabled:
                print(f"SheetFormatter (Targeted): Updating cache for sheet_coord ({actual_sheet_row_idx},{actual_sheet_col_idx}) with style: {cell_style}")
            self._update_cache(actual_sheet_row_idx, actual_sheet_col_idx, cell_style)

def _apply_all_conditions_formatting(self, format_rule, header, values, num_cols, debug_row):
        """Handle all-conditions conditional formatting."""
        for row_idx, row in enumerate(values[1:], 1):
            conditions_met = True
            debug_info = []
    
            # Check all conditions
            for condition in format_rule['conditions']:
                column_name = condition['column']
                if column_name in header:
                    col_index = header.index(column_name)
                    cell_value = row[col_index]
                    condition_met = condition['condition'](cell_value)
                    conditions_met = conditions_met and condition_met
    
                    if self.debug_enabled and row_idx == debug_row:
                        debug_info.append(f"Column '{column_name}': value={cell_value}, condition met={condition_met}")
    
            # Apply formatting if all conditions are met
            if conditions_met:
                format_style = format_rule['format']
                
                if format_rule.get('entire_row', False):
                    # Apply to entire row
                    for col_idx in range(num_cols):
                        current_format = self.formatting_cache[row_idx][col_idx].copy()
                        merged_format = self._merge_formatting(current_format, format_style)
                        self._update_cache(row_idx, col_idx, merged_format)
                else:
                    # Apply only to specified columns
                    for condition in format_rule['conditions']:
                        column_name = condition['column']
                        if column_name in header:
                            col_idx = header.index(column_name)
                            current_format = self.formatting_cache[row_idx][col_idx].copy()
                            merged_format = self._merge_formatting(current_format, format_style)
                            self._update_cache(row_idx, col_idx, merged_format)
    
                # Apply to extra columns if specified
                if format_rule.get('extra_columns'):
                    for extra_col in format_rule['extra_columns']:
                        if extra_col in header:
                            extra_col_idx = header.index(extra_col)
                            extra_format = self.formatting_cache[row_idx][extra_col_idx].copy()
                            merged_extra_format = self._merge_formatting(extra_format, format_style)
                            self._update_cache(row_idx, extra_col_idx, merged_extra_format)
    
                # Debug output for first row
                if self.debug_enabled and row_idx == debug_row:
                    print("\nAll conditions format applied:")
                    for debug_line in debug_info:
                        print(debug_line)
                    print(f"Format applied: {format_style}")
