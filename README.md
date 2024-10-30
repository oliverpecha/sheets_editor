# Sheets Exporter

A lightweight Python package for easily exporting and formatting tabular data to Google Sheets. This package simplifies the process of creating, updating, and formatting Google Sheets programmatically, with built-in support for versioning and consistent formatting.

## Key Features

- 📊 Export any tabular data to Google Sheets
- 🎨 Automatic formatting with customizable options:
  - Alternating row colors
  - Custom row heights
  - Header formatting
- 📂 Version control for sheets
- 🔄 Automatic sheet management (create/update)
- 📧 Configurable sharing options
- 🔐 Secure credentials handling

## Installation

```bash
pip install git+https://github.com/yourusername/sheets_exporter.git

Quick Start

from sheets_exporter import SheetsExporter, SheetConfig
from google.colab import userdata  # For Google Colab
import json

# Configure
config = SheetConfig(
    file_name="MyData",
    share_with=['user@example.com'],
    ignore_columns=['internal_id']
)

# Initialize exporter (using secure credentials)
credentials = json.loads(userdata.get('Service_account'))  # For Google Colab
exporter = SheetsExporter(credentials, config)

# Sample data
data = [
    {'name': 'John', 'age': 30, 'city': 'New York'},
    {'name': 'Jane', 'age': 25, 'city': 'London'}
]

# Export
exporter.export_table(
    data=data,
    version="v1",
    sheet_name="Sheet1"
)

Credentials Security

⚠️ Important: Never commit credentials to GitHub!

This package is designed to work securely with Google Service Account credentials. There are several safe ways to handle credentials:

    Google Colab (Recommended):
        Store credentials in Colab's secure storage
        Access using google.colab.userdata
        Credentials never appear in your notebook or repository

    Environment Variables:

    import os
    import json
    credentials = json.loads(os.environ['GOOGLE_SHEETS_CREDENTIALS'])

    Configuration File (local development only):
        Store credentials in a local credentials.json
        Add credentials.json to .gitignore
        Load credentials from file

Configuration Options

config = SheetConfig(
    file_name="MyData",                    # Base name for sheets
    ignore_columns=['internal_id'],        # Columns to exclude
    share_with=['user@example.com'],       # Auto-share with these emails
    alternate_row_color={                  # Custom row coloring
        "red": 0.9,
        "green": 0.9,
        "blue": 1.0
    },
    track_links=True                       # Track sheet URLs
)

Best Practices

    Credentials Management:
        Never store credentials in code
        Use secure credential storage
        Add credential files to .gitignore

    Version Control:
        Use meaningful version names
        Keep track of sheet versions
        Document major changes

    Data Handling:
        Validate data before export
        Handle empty datasets gracefully
        Consider column ordering

Common Use Cases

    Data Reports:

exporter.export_table(
    data=monthly_report,
    version="v1",
    sheet_name="June2023"
)

Multiple Sheets:

    exporter.export_table(
        data=dataset1,
        version="v1",
        sheet_name="Revenue"
    )
    exporter.export_table(
        data=dataset2,
        version="v1",
        sheet_name="Expenses"
    )

Requirements

    Python 3.7+
    gspread
    google-auth
    pyyaml

Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
License

This project is licensed under the MIT License - see the LICENSE file for details.
