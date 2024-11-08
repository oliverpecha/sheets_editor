from typing import Any, Dict, List

class SheetFormatter:
    @staticmethod
    def format_worksheet(worksheet: Any, formatting_config: Dict = None, conditional_formats: List[Dict] = None) -> None:
        """Applies formatting to the worksheet.

        Args:
            worksheet: The gspread worksheet object.
            formatting_config: A dictionary with absolute formatting settings.
            conditional_formats: A list of dictionaries with conditional formatting settings.
        """
        try:
            if not formatting_config and not conditional_formats:
                return  # Exit if no formatting is required

            print(f"Worksheet object: {worksheet}")
            if formatting_config:
                print("Formatting dictionary:", formatting_config)
                print("Bold rows config:", formatting_config.get('bold_rows'))

            values = worksheet.get_all_values()
            if not values:
                return  # Exit if the sheet is empty

            sheet_id = worksheet._properties['sheetId']
            num_rows = len(values)
            num_cols = len(values[0])

            requests = []

            # Absolute Formatting
            if formatting_config:
                if formatting_config.get('alternate_rows'):
                    for row in range(2, num_rows + 1, 2):  # Alternate rows starting from row 2
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
                            requests.append(SheetFormatter._create_request(row - 1, num_cols, sheet_id, formatting_config['background_color'], True, 0)
                        )  # entire_row=True, col_index=0 for the whole row


                if 'bold_rows' in formatting_config:
                    for row in formatting_config['bold_rows']:
                        requests.append(SheetFormatter._create_request(row -1, num_cols, sheet_id, {'textFormat': {'bold': True}}, True, 0))


            # Conditional Formatting
            if conditional_formats:
                header = values[0]
                for cond_format in conditional_formats:
                    column_name = cond_format.get('column')
                    if not column_name:
                        print("Missing 'column' key in conditional formatting.")
                        continue

                    values_config = cond_format.get('values')
                    condition_func = cond_format.get('condition')
                    entire_row = cond_format.get('entire_row', False)
                    format_style = cond_format.get('format')

                    try:
                        col_index = header.index(column_name)
                    except ValueError:
                        print(f"Column '{column_name}' not found for conditional formatting.")
                        continue

                    for i, row in enumerate(values[1:], 1):
                        cell_value = row[col_index]
                        if values_config and cell_value in values_config:
                            requests.append(SheetFormatter._create_request(i, num_cols, sheet_id, values_config[cell_value], entire_row, col_index))

                        elif condition_func and condition_func(cell_value):  # Corrected condition
                            requests.append(SheetFormatter._create_request(i, num_cols, sheet_id, format_style, entire_row, col_index))


            # Batch Update (execute all requests)
            if requests:
                worksheet.spreadsheet.batch_update({"requests": requests})

        except Exception as e:
            print(f"Error in formatting: {e}")
            raise
