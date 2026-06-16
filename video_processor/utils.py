"""
Shared utilities for video processing: temp storage, chunk uploads, cleanup.
"""
import os
import uuid
import time
import json
import logging
import shutil
from pathlib import Path
from django.conf import settings

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════
# DISK SPACE HELPERS
# ═══════════════════════════════════════════════════════════════

def get_disk_free_mb(path=None):
    """Get free disk space in MB."""
    if path is None:
        path = os.path.dirname(os.path.abspath(__file__))
    try:
        stat = shutil.disk_usage(path)
        return stat.free / (1024 * 1024)  # Convert bytes to MB
    except Exception as e:
        logger.error(f"Failed to get disk space: {e}")
        return 0

# ═══════════════════════════════════════════════════════════════
# PATHS (from settings)
# ═══════════════════════════════════════════════════════════════

def get_uploads_dir():
    return getattr(settings, 'VIDEO_UPLOADS_DIR', os.path.join(settings.BASE_DIR, 'temp_media', 'video', 'uploads'))

def get_outputs_dir():
    return getattr(settings, 'VIDEO_OUTPUTS_DIR', os.path.join(settings.BASE_DIR, 'temp_media', 'video', 'outputs'))

def get_chunks_dir():
    return getattr(settings, 'VIDEO_CHUNKS_DIR', os.path.join(settings.BASE_DIR, 'temp_media', 'video', 'chunks'))

def ensure_dirs():
    os.makedirs(get_uploads_dir(), exist_ok=True)
    os.makedirs(get_outputs_dir(), exist_ok=True)
    os.makedirs(get_chunks_dir(), exist_ok=True)

# ═══════════════════════════════════════════════════════════════
# FILE UPLOAD HELPERS
# ═══════════════════════════════════════════════════════════════

def save_upload(uploaded_file):
    """Save uploaded file to temp uploads directory.
    
    Raises OSError if disk space is insufficient.
    """
    ensure_dirs()
    
    # Check available disk space (require at least 100MB free)
    uploads_dir = get_uploads_dir()
    free_mb = get_disk_free_mb(uploads_dir)
    file_size_mb = uploaded_file.size / (1024 * 1024)
    
    if free_mb < max(file_size_mb + 50, 100):  # Need file size + 50MB buffer, minimum 100MB
        raise OSError(f"[Errno 28] No space left on device (free: {free_mb:.1f}MB, need: {file_size_mb:.1f}MB)")
    
    ext = os.path.splitext(uploaded_file.name)[1]
    file_path = os.path.join(uploads_dir, f"{uuid.uuid4().hex}{ext}")
    
    try:
        with open(file_path, 'wb+') as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)
    except OSError as e:
        # Clean up partial file on failure
        try:
            os.remove(file_path)
        except OSError:
            pass
        raise
    
    return file_path


def safe_remove(path):
    """Safely remove a file, ignoring errors."""
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass


def cleanup_old_files(max_age_seconds=600):
    """Delete files older than max_age_seconds from all video temp dirs.
    
    Returns tuple of (files_removed, space_freed_mb)
    """
    now = time.time()
    removed = 0
    space_freed = 0
    
    for root_dir in [get_uploads_dir(), get_outputs_dir(), get_chunks_dir()]:
        if not os.path.exists(root_dir):
            continue
        try:
            for entry in os.scandir(root_dir):
                if entry.is_file(follow_symlinks=False):
                    if entry.stat().st_mtime < now - max_age_seconds:
                        try:
                            size_bytes = entry.stat().st_size
                            os.remove(entry.path)
                            removed += 1
                            space_freed += size_bytes / (1024 * 1024)  # Convert to MB
                            logger.debug(f"Cleaned: {entry.path} ({size_bytes/1024:.1f}KB)")
                        except OSError as e:
                            logger.warning(f"Failed to remove {entry.path}: {e}")
                elif entry.is_dir(follow_symlinks=False):
                    try:
                        # Try to remove empty subdirectories
                        os.rmdir(entry.path)
                        logger.debug(f"Removed empty dir: {entry.path}")
                    except OSError:
                        # Directory not empty, skip
                        pass
        except OSError as e:
            logger.error(f"Cleanup error in {root_dir}: {e}")
    
    logger.info(f"Cleanup: Removed {removed} files, freed {space_freed:.1f}MB")
    return removed, space_freed

# ═══════════════════════════════════════════════════════════════
# CHUNK UPLOAD SYSTEM
# ═══════════════════════════════════════════════════════════════

def save_chunk(upload_id, chunk_index, chunk_file):
    """Save an uploaded chunk to disk.
    
    Raises OSError with appropriate message if disk space is insufficient.
    """
    chunk_dir = os.path.join(get_chunks_dir(), upload_id)
    os.makedirs(chunk_dir, exist_ok=True)
    chunk_path = os.path.join(chunk_dir, f"{chunk_index}.part")
    
    # Check available disk space (require at least 50MB free)
    free_mb = get_disk_free_mb(chunk_dir)
    chunk_size_mb = chunk_file.size / (1024 * 1024)  # Use .size attribute instead of iterating
    
    if free_mb < max(chunk_size_mb + 10, 50):  # Need chunk size + 10MB buffer, minimum 50MB
        raise OSError(f"[Errno 28] No space left on device (free: {free_mb:.1f}MB, need: {chunk_size_mb:.1f}MB)")
    
    try:
        with open(chunk_path, 'wb+') as destination:
            for chunk in chunk_file.chunks():
                destination.write(chunk)
    except OSError as e:
        # Clean up partial chunk on failure
        try:
            os.remove(chunk_path)
        except OSError:
            pass
        raise
    
    logger.debug(f"Saved chunk {chunk_index} for upload {upload_id} ({chunk_size_mb:.1f}MB)")
    return chunk_path


def assemble_chunks(upload_id, total_chunks, original_filename):
    """Assemble chunks into final file and return path."""
    chunk_dir = os.path.join(get_chunks_dir(), upload_id)
    ext = os.path.splitext(original_filename)[1]
    final_path = os.path.join(get_uploads_dir(), f"{upload_id}{ext}")

    with open(final_path, 'wb') as outfile:
        for i in range(total_chunks):
            chunk_path = os.path.join(chunk_dir, f"{i}.part")
            if not os.path.exists(chunk_path):
                raise FileNotFoundError(f"Missing chunk {i} for upload {upload_id}")
            with open(chunk_path, 'rb') as infile:
                outfile.write(infile.read())

    # Cleanup chunks
    try:
        for i in range(total_chunks):
            safe_remove(os.path.join(chunk_dir, f"{i}.part"))
        os.rmdir(chunk_dir)
    except OSError:
        pass

    return final_path


def get_chunk_status(upload_id, total_chunks):
    """Return which chunks have been received."""
    chunk_dir = os.path.join(get_chunks_dir(), upload_id)
    received = set()
    if os.path.exists(chunk_dir):
        try:
            for entry in os.scandir(chunk_dir):
                if entry.is_file(follow_symlinks=False) and entry.name.endswith('.part'):
                    try:
                        idx = int(entry.name.replace('.part', ''))
                        received.add(idx)
                        logger.debug(f"Found chunk {idx} in {upload_id}")
                    except ValueError:
                        logger.warning(f"Skipped invalid chunk file: {entry.name}")
        except OSError as e:
            logger.error(f"Error scanning chunk directory {chunk_dir}: {e}")
    
    missing = sorted(set(range(total_chunks)) - received)
    complete = len(received) == total_chunks
    
    logger.debug(f"Chunk status for {upload_id}: received={sorted(received)} missing={missing} complete={complete}")
    
    return {
        'upload_id': upload_id,
        'total_chunks': total_chunks,
        'received_chunks': sorted(received),
        'missing_chunks': missing,
        'complete': complete
    }

# ═══════════════════════════════════════════════════════════════
# RESPONSE HELPERS
# ═══════════════════════════════════════════════════════════════

class FileCleanupResponse:
    """
    Wrapper for Django FileResponse that auto-deletes the file after streaming.
    Usage: return FileCleanupResponse(file_path, content_type, filename)
    """
    def __init__(self, file_path, content_type=None, filename=None):
        self.file_path = file_path
        self.content_type = content_type or 'application/octet-stream'
        self.filename = filename or os.path.basename(file_path)

    def build_response(self):
        from django.http import FileResponse
        response = FileResponse(open(self.file_path, 'rb'), content_type=self.content_type)
        response['Content-Disposition'] = f'attachment; filename="{self.filename}"'
        # Use close callback to delete after streaming
        original_close = response.close

        def close_and_delete():
            original_close()
            safe_remove(self.file_path)

        response.close = close_and_delete
        return response

