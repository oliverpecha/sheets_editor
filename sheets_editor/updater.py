from enum import Enum
import gspread

class Alignment(Enum):
    """
    Enumeration for cell alignment options.
    """
    LEFT = "LEFT"
    RIGHT = "RIGHT"
    CENTER = "CENTER"
    FLOAT = "FLOAT"  # Custom option for centering both horizontally and vertically

class SheetUpdater:
    def __init__(self, worksheet):
        """Initialize with the Google Sheets worksheet instance."""
        self.sheet = worksheet

    def update_image_formulas(self, cell_updates: list[tuple[int, int, str]]):
        """
        Updates the cells on the spreadsheet with the provided formulas.

        Args:
            cell_updates (list[tuple[int, int, str]]): A list of tuples, where each tuple
                                                    contains (row, col, formula_string).
        Example:
            cell_updates = [(1, 1, '=IMAGE("https://example.com/image.png")'),
                            (2, 1, '=IMAGE("https://example.com/another_image.jpg")')]
        """

        cells_to_update = []  # List to store the Cell objects to update

        for row, col, formula_string in cell_updates:
              try:
                  cells_to_update.append(gspread.Cell(row, col, formula_string)) #Create cell object from data provided in input
              except Exception as e:
                  print(f"Error updating cell ({row}, {col}): {e}")
                  continue

        try:
            if cells_to_update:
                 self.sheet.update_cells(cells_to_update, value_input_option='USER_ENTERED')  # Update all cells in one go. USER_ENTERED makes sure that the formula string is parsed as a formula.
            else:
                print ("No cells to update.")

        except Exception as e:
                print(f"Error during batch update: {e}")

    def update_image_formulas_with_formatting(
        self,
        cell_updates: list[tuple[int, int, str]],
        column_width: int = None,
        row_height: int = None,
        alignment: Alignment = None,
    ):
        """
        Updates cells with image formulas and optionally sets column width, row height, and alignment.

        Args:
            cell_updates (list[tuple[int, int, str]]): A list of tuples, where each tuple
                                                        contains (row, col, formula_string).
            column_width (int, optional): The desired width of the column in pixels. Defaults to None.
            row_height (int, optional): The desired height of the row in pixels. Defaults to None.
            alignment (Alignment, optional): The desired alignment (LEFT, RIGHT, CENTER, or FLOAT). Defaults to None.
                                              FLOAT sets both horizontal and vertical alignment to center.
        Example:
            cell_updates = [(1, 1, '=IMAGE("https://example.com/image.png")'),
                            (2, 1, '=IMAGE("https://example.com/another_image.jpg")')]
            updater.update_image_formulas_extended(cell_updates, column_width=200, row_height=100, alignment=Alignment.CENTER)
            #sets images to cells (1,1) and (2,1), sets column width for column 1 to 200px, row height for rows 1 and 2 to 100 px and horizontal alignment for cells to center.
        """
        self.update_image_formulas(cell_updates) #update images first

        if column_width is not None:
            self.set_column_width(cell_updates, column_width)

        if row_height is not None:
            self.set_row_height(cell_updates, row_height)

        if alignment is not None:
            self.set_cell_alignment(cell_updates, alignment)

    def set_column_width(self, cell_updates, column_width: int):
         """Sets the width of the columns corresponding to the updated cells.

         Args:
            cell_updates (list[tuple[int, int, str]]): A list of tuples, where each tuple
                                                        contains (row, col, formula_string).
            column_width (int): The desired width of the column in pixels.
         """
         columns_to_update = set()
         for _, col, _ in cell_updates:
             columns_to_update.add(col)

         body = {
              "requests": [
                 {
                    "updateDimensionProperties": {
                       "range": {
                          "sheetId": self.sheet.id,
                          "dimension": "COLUMNS",
                          "startIndex": col - 1,  # API indices are zero-based
                          "endIndex": col
                       },
                       "properties": {
                          "pixelSize": column_width
                       },
                       "fields": "pixelSize"
                    }
                 }
               for col in columns_to_update
             ]
          }

         try:
            self.sheet.spreadsheet.batch_update(body)
         except Exception as e:
            print(f"Error setting column width: {e}")

    def set_row_height(self, cell_updates, row_height: int):
          """Sets the height of the rows corresponding to the updated cells.

          Args:
              cell_updates (list[tuple[int, int, str]]): A list of tuples, where each tuple
                                                          contains (row, col, formula_string).
              row_height (int): The desired height of the row in pixels.
          """
          rows_to_update = set()
          for row, _, _ in cell_updates:
              rows_to_update.add(row)

          body = {
              "requests": [
                 {
                    "updateDimensionProperties": {
                       "range": {
                          "sheetId": self.sheet.id,
                          "dimension": "ROWS",
                          "startIndex": row - 1,  # API indices are zero-based
                          "endIndex": row
                       },
                       "properties": {
                          "pixelSize": row_height
                       },
                       "fields": "pixelSize"
                    }
                 } for row in rows_to_update
               ]

          }

          try:
            self.sheet.spreadsheet.batch_update(body)
          except Exception as e:
            print(f"Error setting row height: {e}")

    def set_cell_alignment(self, cell_updates, alignment: Alignment):
        """
        Sets the alignment of the cells corresponding to the updated formulas.
    
        Args:
            cell_updates (list[tuple[int, int, str]]): A list of tuples, where each tuple
                                                       contains (row, col, formula_string).
            alignment (Alignment): The desired alignment (LEFT, RIGHT, CENTER, or FLOAT).
                                   FLOAT sets both horizontal and vertical alignment to center.
        """
        alignment_value = alignment.value  # Get the enum value (e.g., 'LEFT', 'RIGHT', etc.)
        
        body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": self.sheet.id,
                            "startRowIndex": row - 1,  # API indices are zero-based
                            "endRowIndex": row,
                            "startColumnIndex": col - 1,
                            "endColumnIndex": col,
                        },
                        "fields": "userEnteredFormat.horizontalAlignment, userEnteredFormat.verticalAlignment",
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "CENTER" if alignment_value == "FLOAT" else alignment_value,
                                "verticalAlignment": "MIDDLE" if alignment_value == "FLOAT" else None
                            }
                        },
                    }
                }
                for row, col, _ in cell_updates
            ]
        }
        
        try:
            self.sheet.spreadsheet.batch_update(body)
        except Exception as e:
            print(f"Error setting cell alignment: {e}")

def parse_alignment(self, alignment_str: str) -> Alignment:
        """
        Converts the alignment string to an Alignment enum.
        Args:
            alignment_str (str): The alignment string to parse.
        Returns:
            Alignment: The corresponding Alignment enum value.
        Raises:
            ValueError: If the provided alignment string is not valid.
        """
        try:
            return Alignment[alignment_str.upper()]
        except KeyError:
            raise ValueError(f"Invalid alignment value: {alignment_str}. "
                             "Valid options are: 'LEFT', 'RIGHT', 'CENTER', 'FLOAT'.")
