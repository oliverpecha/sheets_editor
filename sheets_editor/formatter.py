from typing import Any, Dict
#Version: "publish to production"

class SheetFormatter:
    def __init__(self):
        """Initialize the formatting cache."""
        self.formatting_cache = {}
        self.debug_enabled = True
        
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

    def format_worksheet(self, worksheet, formatting_config=None, conditional_formats=None):
        """
        Main public method to format a worksheet with both absolute and conditional formatting.
        
        Args:
            worksheet: The worksheet object to format
            formatting_config: Optional dictionary containing absolute formatting rules
            conditional_formats: Optional list of conditional formatting rules
        """
        if self.debug_enabled:
            print(f"\nFormatting worksheet: {worksheet.title}")
            print(f"Absolute formatting config: {formatting_config}")
            print(f"Number of conditional formats: {len(conditional_formats) if conditional_formats else 0}")
    
        # Get all values from worksheet
        values = worksheet.get_all_values()
        if not values:
            print("No values found in worksheet")
            return
    
        # Get worksheet properties
        num_rows = len(values)
        num_cols = len(values[0]) if values else 0
        sheet_id = worksheet.id
    
        try:
            # Initialize the formatting cache
            self._initialize_cache(num_rows, num_cols)
    
            # Apply absolute formatting if configured
            if formatting_config:
                self._apply_absolute_formatting(formatting_config, sheet_id, num_rows, num_cols)
    
            # Apply conditional formatting if configured
            if conditional_formats:
                self._apply_conditional_formatting(conditional_formats, sheet_id, values)
    
            # Generate formatting requests
            requests = self._generate_requests_from_cache(sheet_id, num_cols)
            
            if self.debug_enabled:
                print(f"\nGenerated {len(requests)} formatting requests")
                if requests:
                    print("Sample request:", requests[0])
    
            # Apply the formatting requests directly
            if requests:
                try:
                    worksheet.spreadsheet.batch_update({"requests": requests})
                    if self.debug_enabled:
                        print("✅ Successfully applied formatting requests")
                except Exception as e:
                    print(f"❌ Error applying formatting requests: {str(e)}")
                    raise
    
        except Exception as e:
            print(f"Error during formatting: {str(e)}")
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
