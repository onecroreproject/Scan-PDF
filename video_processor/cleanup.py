"""
Temporary file cleanup system for video processing.
Automatically deletes old files to prevent storage buildup.
"""
import os
import time
import threading
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════

DEFAULT_MAX_AGE_SECONDS = 600  # 10 minutes
CLEANUP_INTERVAL_SECONDS = 300  # 5 minutes

# ═══════════════════════════════════════════════════════════════
# CLEANUP FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def cleanup_directory(directory, max_age_seconds=DEFAULT_MAX_AGE_SECONDS):
    """Remove files older than max_age_seconds from a directory."""
    if not os.path.exists(directory):
        return 0

    now = time.time()
    removed = 0
    try:
        for entry in os.scandir(directory):
            if entry.is_file():
                try:
                    if entry.stat().st_mtime < now - max_age_seconds:
                        os.remove(entry.path)
                        removed += 1
                except OSError:
                    pass
            elif entry.is_dir():
                # Recursively clean subdirectories
                sub_removed = cleanup_directory(entry.path, max_age_seconds)
                removed += sub_removed
                # Remove empty dirs
                try:
                    os.rmdir(entry.path)
                except OSError:
                    pass
    except OSError as e:
        logger.error(f"Cleanup error in {directory}: {e}")

    return removed


def cleanup_temp_dirs(temp_dirs, max_age_seconds=DEFAULT_MAX_AGE_SECONDS):
    """Clean multiple temp directories."""
    total_removed = 0
    for directory in temp_dirs:
        count = cleanup_directory(directory, max_age_seconds)
        total_removed += count
        if count > 0:
            logger.info(f"Cleaned {count} old files from {directory}")
    return total_removed


# ═══════════════════════════════════════════════════════════════
# SCHEDULER
# ═══════════════════════════════════════════════════════════════

class CleanupScheduler:
    """
    Background thread that periodically cleans up temp directories.
    Usage:
        scheduler = CleanupScheduler([dir1, dir2], interval=300)
        scheduler.start()
    """
    def __init__(self, directories, interval=CLEANUP_INTERVAL_SECONDS, max_age=DEFAULT_MAX_AGE_SECONDS):
        self.directories = directories
        self.interval = interval
        self.max_age = max_age
        self._stop_event = threading.Event()
        self._thread = None

    def start(self):
        """Start the cleanup thread."""
        if self._thread is not None and self._thread.is_alive():
            logger.warning("CleanupScheduler already running")
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        logger.info("CleanupScheduler started")

    def stop(self):
        """Stop the cleanup thread."""
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=5)
        logger.info("CleanupScheduler stopped")

    def _run(self):
        while not self._stop_event.is_set():
            try:
                cleanup_temp_dirs(self.directories, self.max_age)
            except Exception as e:
                logger.error(f"Scheduled cleanup error: {e}")
            # Wait for interval or until stopped
            self._stop_event.wait(self.interval)


# ═══════════════════════════════════════════════════════════════
# DJANGO MANAGEMENT COMMAND HELPERS
# ═══════════════════════════════════════════════════════════════

def immediate_cleanup_all(base_temp_dir, max_age=DEFAULT_MAX_AGE_SECONDS):
    """Force immediate cleanup of all temp subdirectories."""
    subdirs = []
    if os.path.exists(base_temp_dir):
        for entry in os.scandir(base_temp_dir):
            if entry.is_dir():
                subdirs.append(entry.path)
    return cleanup_temp_dirs(subdirs, max_age)


def ensure_temp_dirs(dirs):
    """Create temp directories if they don't exist."""
    for d in dirs:
        os.makedirs(d, exist_ok=True)
