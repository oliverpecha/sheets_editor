# Sheets Editor

A Python package for managing Google Sheets with features for exporting, formatting, and manipulating spreadsheet data. Currently implements exporting functionality with more editing features coming soon.

## Features

### Current Features
- 📤 **Data Export**
  - Export any tabular data to Google Sheets
  - Automatic sheet creation and management
  - Version control for sheets
  - Configurable column filtering

- 🎨 **Formatting**
  - Automatic alternate row coloring
  - Customizable row heights
  - Header formatting
  - Custom background colors

- 🔐 **Security**
  - Secure credential handling
  - Configurable sharing options
  - Google Workspace integration

### Coming Soon
- 📝 Data editing and cell manipulation
- 📊 Formula management
- 🔄 Data synchronization
- 📱 Mobile-friendly formatting options

## Installation

```bash
pip install git+https://github.com/oliverpecha/sheets_editor.git

Quick Start

from sheets_editor import SheetsExporter, SheetConfig
from google.colab import userdata  # For Google Colab
import json

# Configure
config = SheetConfig(
    file_name="MyData",
    share_with=['user@example.com'],
    ignore_columns=['internal_id']
)

# Initialize with secure credentials
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

Credentials Security

⚠️ Important: Never commit credentials to GitHub!

This package is designed to work securely with Google Service Account credentials. Recommended ways to handle credentials:

    Google Colab (Recommended):

from google.colab import userdata
credentials = json.loads(userdata.get('Service_account'))

Environment Variables:

    import os
    import json
    credentials = json.loads(os.environ['GOOGLE_SHEETS_CREDENTIALS'])

    Local Development:
        Store credentials in credentials.json
        Add credentials.json to .gitignore
        Load credentials from file

Usage Examples
Basic Export

exporter.export_table(
    data=data,
    version="v1",
    sheet_name="Sheet1"
)

With Custom Formatting

config = SheetConfig(
    file_name="MyData",
    alternate_row_color={
        "red": 0.9,
        "green": 0.9,
        "blue": 1.0
    }
)

Multiple Sheets

# Export different datasets to different sheets
exporter.export_table(data=dataset1, version="v1", sheet_name="Revenue")
exporter.export_table(data=dataset2, version="v1", sheet_name="Expenses")

Best Practices

    Data Preparation
        Validate data before export
        Clean column names
        Handle missing values

    Version Control
        Use meaningful version names
        Document major changes
        Keep track of sheet versions

    Security
        Secure credential storage
        Minimal sharing permissions
        Regular access review

Requirements

    Python 3.7+
    gspread
    google-auth
    pyyaml

Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
Development Setup

    Clone the repository
    Install dependencies: pip install -r requirements.txt
    Create a feature branch
    Submit a pull request

License

This project is licensed under the MIT License - see the LICENSE file for details.
Support

    Report issues on GitHub
    Submit feature requests
    Contact: [Your Contact Information]

