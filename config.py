from typing import Dict, List, Optional

class SheetConfig:
    def __init__(self, 
                 file_name: str,
                 ignore_columns: Optional[List[str]] = None,
                 share_with: Optional[List[str]] = None,
                 alternate_row_color: Optional[Dict] = None,
                 track_links: bool = True):
        self.file_name = file_name
        self.ignore_columns = ignore_columns or []
        self.share_with = share_with or []
        self.alternate_row_color = alternate_row_color or {
            "red": 0.9,
            "green": 0.9,
            "blue": 1.0
        }
        self.track_links = track_links
