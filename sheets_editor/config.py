from typing import List, Optional

class SheetConfig:
    def __init__(self, 
                 file_name: str,
                 ignore_columns: Optional[List[str]] = None,
                 share_with: Optional[List[str]] = None,
                 track_links: bool = True):  

        self.file_name = file_name
        self.ignore_columns = ignore_columns or []
        self.share_with = share_with or []
        self.track_links = track_links
