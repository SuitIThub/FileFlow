"""File system watcher for monitoring new files"""
import time
import threading
from watchdog.events import FileSystemEventHandler


class FileWatcher(FileSystemEventHandler):
    """Handles file system events"""
    
    def __init__(self, app):
        self.app = app
        self.last_check_time = time.time()
    
    def on_created(self, event):
        if not event.is_directory:
            # Add a small delay to ensure file is fully written
            threading.Timer(0.5, self._process_new_file, args=[event.src_path]).start()
    
    def _process_new_file(self, file_path):
        # Wait a bit longer to ensure file is fully written
        time.sleep(0.5)
        
        # Check if file is still being written by trying to open it
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # Try to open the file to ensure it's not being written
                with open(file_path, 'rb') as f:
                    f.read(1)  # Try to read a byte
                break
            except (PermissionError, IOError):
                if attempt < max_retries - 1:
                    time.sleep(1.0)  # Wait longer between retries
                else:
                    return  # Skip this file if we can't access it
        
        self.app.add_tracked_file(file_path)

