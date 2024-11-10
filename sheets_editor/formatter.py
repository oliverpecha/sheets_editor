def _apply_conditional_formatting(self, conditional_formats, sheet_id, num_cols, values):
    """Applies conditional formatting based on the provided conditions."""
    requests = []
    header = values[0]  # Assuming the first row contains headers (column names)

    for cond_format in conditional_formats:
        conditions = cond_format.get('conditions', [])
        format_name = cond_format.get('name', 'Unnamed Format')
        formatting_type = cond_format.get('type', 'case_specific')  # Default to case_specific

        print(f"Processing conditional format '{format_name}' of type '{formatting_type}' with conditions: {conditions}")

        if not conditions:
            print("No conditions provided for conditional formatting.")
            continue

        # Iterate over the data rows (skipping the header row)
        for i, row in enumerate(values[1:], 1):  # Start from the second row (data rows)

            if formatting_type == 'case_specific':
                for index, condition in enumerate(conditions):
                    column_name = condition.get('column')
                    condition_func = condition.get('condition')
                    format_style = cond_format.get('format')[index] if isinstance(cond_format.get('format'), list) else cond_format.get('format')

                    # Ensure the column exists
                    if column_name not in header:
                        print(f"Column '{column_name}' not found in the header.")
                        continue

                    col_index = header.index(column_name)  # Find the column index
                    cell_value = row[col_index]  # Get the cell's value

                    # Check if the condition is met
                    if condition_func(cell_value):
                        print(f"Applying case-specific formatting for '{format_name}' on row {i + 1}: {cell_value}")
                        if cond_format.get('entire_row', False):
                            # Apply to the entire row
                            requests.append(self._create_request(i, num_cols, sheet_id, format_style, True, 0))  # Apply to entire row
                        else:
                            requests.append(self._create_request(i, num_cols, sheet_id, format_style, False, col_index))  # Apply to specific column

            elif formatting_type == 'all_conditions':
                all_conditions_met = True
                for condition in conditions:
                    column_name = condition.get('column')
                    condition_func = condition.get('condition')

                    # Ensure the column exists
                    if column_name not in header:
                        print(f"Column '{column_name}' not found in the header.")
                        all_conditions_met = False
                        break  # Exit the loop since one condition is not met

                    col_index = header.index(column_name)  # Find the column index
                    cell_value = row[col_index]  # Get the cell's value

                    # Check if the condition is met
                    if not condition_func(cell_value):
                        all_conditions_met = False
                        break

                # If all conditions are met, apply the formatting
                if all_conditions_met:
                    print(f"Applying all-conditions formatting for '{format_name}' to row {i + 1}")
                    requests.append(self._create_request(i, num_cols, sheet_id, cond_format.get('format'), True, 0))  # Apply to entire row

    return requests
