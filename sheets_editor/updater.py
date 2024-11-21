class SheetUpdater:
    def __init__(self, worksheet):
        """Initialize with the Google Sheets worksheet instance."""
        self.worksheet = worksheet

    def update_image_formulas(self, image_cells):
        """
        Updates the worksheet with image formulas in the specified cells.
        :param image_cells: List of tuples containing (row, column, image_formula)
        """
        for row, col, image_formula in image_cells:
            # Set the image formula in the appropriate cell
            self.worksheet.update_cell(row, col, image_formula)
