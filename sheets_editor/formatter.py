from typing import Any, Dict, List, Optional, Union

class SheetFormatter:
    def __init__(self):
        pass

    def _create_request(self, row_index: int, num_cols: int, sheet_id: int, format_style: Dict[str, Any], 
                        entire_row: bool, col_index: Optional[int] = None) -> Dict[str, Any]:

        start_col = 0 if entire_row else col_index or 0
        end_col = num_cols if entire_row else (col_index + 1 if col_index is not None else num_cols)

        user_entered_format = {}
        if format_style:
            user_entered_format.update(format_style)

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
                "fields": ",".join(user_entered_format.keys()) if user_entered_format else ""
            }
        }

    def _merge_format_styles(self, existing_style: Dict[str, Any], new_style: Dict[str, Any]) -> Dict[str, Any]:
        merged_style = existing_style.copy()
        merged_style.update(new_style)

        if "textFormat" in new_style:
            merged_style.setdefault("textFormat", {}).update(new_style["textFormat"])

        return merged_style

    def format_worksheet(self, worksheet, formatting_config=None, conditional_formats=None):
        values = worksheet.get_all_values()
        if not values:
            return

        header = values[0]
        num_cols = len(header)
        requests = []
        existing_formats = {}

        if formatting_config:
            requests.extend(self._apply_absolute_formatting(formatting_config, worksheet._properties['sheetId'], len(values), num_cols, values))
            for req in requests:
                if "repeatCell" in req:
                    row_index = req["repeatCell"]["range"]["startRowIndex"]
                    existing_formats[row_index] = req["repeatCell"]["cell"]["userEnteredFormat"]

        if conditional_formats:
            requests.extend(self._apply_conditional_formatting(conditional_formats, worksheet._properties['sheetId'], values, existing_formats))

        if requests:
            worksheet.spreadsheet.batch_update({"requests": requests})

    
    def _apply_absolute_formatting(self, formatting_config, sheet_id, num_rows, num_cols, values):
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
                    requests.append(self._create_request(row - 1, num_cols, sheet_id, formatting_config['background_color'], True, 0))

        if 'bold_rows' in formatting_config:
            for row in formatting_config['bold_rows']:
                requests.append(self._create_request(row - 1, num_cols, sheet_id, {'textFormat': {'bold': True}}, True, 0))
        return requests



    def _apply_conditional_formatting(self, conditional_formats, sheet_id, values, existing_formats):
        """Applies conditional formatting based on provided conditions and merges with existing styles."""
        requests = []
        header = values[0]
        num_cols = len(header)

        for cond_format in conditional_formats:
            conditions = cond_format.get('conditions', [])
            format_name = cond_format.get('name', 'Unnamed Format')
            formatting_type = cond_format.get('type', 'case_specific')
            entire_row = cond_format.get('entire_row', False)
            extra_columns = cond_format.get('extra_columns', [])

            print(f"Processing conditional format '{format_name}' of type '{formatting_type}' with conditions: {conditions}")

            if not conditions:
                print("No conditions provided for conditional formatting.")
                continue

            for i, row in enumerate(values[1:], 1):
                format_style = {}
                col_index = None

                if formatting_type == 'case_specific':
                    matched_case = False
                    for index, condition in enumerate(conditions):
                        if self._check_condition(condition, row, header):
                            if isinstance(cond_format.get('format'), list):
                                format_style = cond_format['format'][index]
                            else:
                                format_style = cond_format['format']
                            matched_case = True
                            break  # Exit after first match
                    if not matched_case:
                        continue

                elif formatting_type == 'all_conditions':
                    if all(self._check_condition(condition, row, header) for condition in conditions):
                        format_style = cond_format['format']
                    else:
                        continue

                if extra_columns and format_style:
                    for col_name in extra_columns:
                        try:
                            col_index = header.index(col_name)
                            merged_style = self._merge_format_styles(existing_formats.get(i, {}), format_style)

                            requests.append(self._create_request(i, num_cols, sheet_id, merged_style, entire_row=False, col_index=col_index))

                        except ValueError:
                            print(f"Column '{col_name}' not found in header.")

                if format_style:
                    merged_style = self._merge_format_styles(existing_formats.get(i, {}), format_style)
                    requests.append(self._create_request(i, num_cols, sheet_id, merged_style, entire_row, col_index))
        return requests

    def _check_condition(self, condition, row, header):
        """Checks a single condition against a row."""
        try:
            col_name = condition['column']
            col_index = header.index(col_name)
            cell_value = row[col_index]
            return condition['condition'](cell_value)
        except ValueError:
            print(f"Column '{condition['column']}' not found in header. Skipping condition.")
            return False
        except Exception as e:
            print(f"Error evaluating condition for column '{condition['column']}': {e}")
            return False
