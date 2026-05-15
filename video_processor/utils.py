import os
import uuid
import tempfile
import time
import shutil
from pathlib import Path
from moviepy.editor import VideoFileClip
from django.conf import settings

# ═══════════════════════════════════════════════════════════════
# STORAGE CONFIGURATION
# ═══════════════════════════════════════════════════════════════

BASE_TEMP_DIR = os.path.join(settings.BASE_DIR, 'temp_media', 'video_converter')
UPLOADS_DIR = os.path.join(BASE_TEMP_DIR, 'uploads')
OUTPUTS_DIR = os.path.join(BASE_TEMP_DIR, 'outputs')

def ensure_dirs():
    """Ensure required directories exist."""
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    os.makedirs(OUTPUTS_DIR, exist_ok=True)

def cleanup_old_files(max_age_seconds=600):
    """Delete files older than max_age_seconds (default 10 mins)."""
    now = time.time()
    for root_dir in [UPLOADS_DIR, OUTPUTS_DIR]:
        if not os.path.exists(root_dir):
            continue
        for filename in os.listdir(root_dir):
            file_path = os.path.join(root_dir, filename)
            try:
                if os.path.getmtime(file_path) < now - max_age_seconds:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
            except Exception as e:
                print(f"Cleanup error: {e}")

def convert_video_moviepy(input_path, output_format):
    """
    Converts a video using ONLY MoviePy methods.
    Returns the path to the converted file.
    """
    ensure_dirs()
    output_filename = f"{uuid.uuid4().hex}.{output_format.lower()}"
    output_path = os.path.join(OUTPUTS_DIR, output_filename)
    
    try:
        # Load clip
        clip = VideoFileClip(input_path)
        
        # Determine codec based on format
        codec = 'libx264'
        if output_format.lower() == 'webm':
            codec = 'libvpx'
        elif output_format.lower() == 'ogv':
            codec = 'libtheora'
        
        # Write file using MoviePy's write_videofile
        # Note: logger=None suppresses console progress bars in production
        clip.write_videofile(
            output_path, 
            codec=codec, 
            audio_codec='aac' if codec != 'libvpx' else 'libvorbis',
            logger=None
        )
        
        # Close clip to release file handle
        clip.close()
        
        return output_path
    except Exception as e:
        if os.path.exists(output_path):
            os.remove(output_path)
        raise e

def save_upload(uploaded_file):
    """Save uploaded file to temp uploads directory."""
    ensure_dirs()
    ext = os.path.splitext(uploaded_file.name)[1]
    file_path = os.path.join(UPLOADS_DIR, f"{uuid.uuid4().hex}{ext}")
    with open(file_path, 'wb+') as destination:
        for chunk in uploaded_file.chunks():
            destination.write(chunk)
    return file_path
