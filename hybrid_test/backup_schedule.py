from datetime import date
from pathlib import Path
from safe_backup import snapshot

class DailyBackup:
    """Invoke from a background worker, not the UI thread."""
    def __init__(self, source, directory, snapshot_fn=snapshot):
        self.source = source
        self.directory = Path(directory)
        self.snapshot_fn = snapshot_fn
        self.last_success = None

    def tick(self, today=None):
        today = today or date.today()
        if self.last_success == today:
            return None
        result = self.snapshot_fn(self.source, self.directory)
        self.last_success = today
        return result
