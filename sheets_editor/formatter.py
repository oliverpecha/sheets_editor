from typing import Any, Dict, List

class SheetFormatter:
    def format_worksheet(worksheet: Any, formatting_config: Dict = None, conditional_formats: List[Dict] = None) -> None:
        """Applies formatting to the worksheet.

        Args:
            worksheet: The gspread worksheet object.
            formatting_config: A dictionary with absolute formatting settings.
            conditional_formats: A list of dictionaries with conditional formatting settings.
        """
        if not formatting_config and not conditional_formats:
            return

        try:
            values = worksheet.get_all_values()
            if not values:
                return

            sheet_id = worksheet._properties['sheetId']
            num_rows = len(values)
            num_cols = len(values[0])
            requests = []

            if formatting_config:
                requests.extend(SheetFormatter._apply_absolute_formatting(formatting_config, sheet_id, num_rows, num_cols, values))

            if conditional_formats:
                requests.extend(SheetFormatter._apply_conditional_formatting(conditional_formats, sheet_id, num_cols, values))

            if requests:
                worksheet.spreadsheet.batch_update({"requests": requests})

        except Exception as e:
            print(f"Error in formatting: {e}")
            raise

            # Batch Update (execute all requests)
            if requests:
                worksheet.spreadsheet.batch_update({"requests": requests})

        except Exception as e:
            print(f"Error in formatting: {e}")
            raise

    def _apply_absolute_formatting(formatting_config, sheet_id, num_rows, num_cols, values):
        requests = []
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
                    requests.append(SheetFormatter._create_request(row-1, num_cols, sheet_id, formatting_config['background_color'], True, 0))

        if 'bold_rows' in formatting_config:
             for row in formatting_config['bold_rows']:
                requests.append(SheetFormatter._create_request(row - 1, num_cols, sheet_id, {'textFormat': {'bold': True}}, True, 0))
        return requests

    def _apply_conditional_formatting(conditional_formats, sheet_id, num_cols, values):
        requests = []
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
        return requests

    def _create_request(row_index, num_cols, sheet_id, format_style, entire_row, col_index):
        """Creates a batch update request for formatting a cell or row."""

        start_col = 0 if entire_row else col_index
        end_col = num_cols if entire_row else col_index + 1
        return {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_index,
                    "endRowIndex": row_index+1,
                    "startColumnIndex": start_col,
                    "endColumnIndex": end_col
                },
                "cell": {
                    "userEnteredFormat": format_style
                },
                "fields": "userEnteredFormat" #Use all fields in format_style
            }

        }
