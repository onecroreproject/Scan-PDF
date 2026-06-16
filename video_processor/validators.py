"""
Security validators for video processing uploads.
"""
import os
import mimetypes
from django.core.exceptions import ValidationError
from django.conf import settings

try:
    import magic
except ImportError:
    magic = None  # Fallback on Windows without libmagic; install libmagic1 on Linux

# ═══════════════════════════════════════════════════════════════
# ALLOWED FORMATS
# ═══════════════════════════════════════════════════════════════

ALLOWED_VIDEO_EXTENSIONS = {
    '.mp4', '.avi', '.mov', '.mkv', '.webm', '.m4v', '.flv', '.wmv'
}

ALLOWED_IMAGE_EXTENSIONS = {
    '.jpg', '.jpeg', '.png', '.bmp', '.webp', '.gif', '.tiff'
}

ALLOWED_AUDIO_EXTENSIONS = {
    '.mp3', '.wav', '.aac', '.flac', '.ogg', '.m4a', '.wma'
}

ALLOWED_SUBTITLE_EXTENSIONS = {
    '.srt', '.vtt', '.ass', '.ssa'
}

VIDEO_MIME_TYPES = {
    'video/mp4', 'video/x-msvideo', 'video/quicktime',
    'video/x-matroska', 'video/webm', 'video/x-flv',
    'video/x-ms-wmv', 'video/m4v'
}

IMAGE_MIME_TYPES = {
    'image/jpeg', 'image/png', 'image/bmp', 'image/webp',
    'image/gif', 'image/tiff'
}

AUDIO_MIME_TYPES = {
    'audio/mpeg', 'audio/wav', 'audio/x-wav', 'audio/aac',
    'audio/flac', 'audio/ogg', 'audio/x-m4a', 'audio/mp4',
    'audio/x-ms-wma'
}

# ═══════════════════════════════════════════════════════════════
# VALIDATION FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def validate_upload_size(file_obj, max_mb=2048):
    """Validate file size against max MB limit."""
    max_bytes = max_mb * 1024 * 1024
    if hasattr(file_obj, 'size') and file_obj.size > max_bytes:
        raise ValidationError(
            f"File too large. Maximum allowed size is {max_mb}MB. "
            f"Your file is {file_obj.size / (1024*1024):.1f}MB."
        )
    return True


def validate_extension(filename, allowed_extensions):
    """Validate file extension is in allowed set."""
    ext = os.path.splitext(filename)[1].lower()
    if ext not in allowed_extensions:
        raise ValidationError(
            f"Unsupported file format: {ext}. Allowed: {', '.join(sorted(allowed_extensions))}"
        )
    return ext


def validate_mime_type(file_path, allowed_mime_types):
    """Validate actual file content using python-magic."""
    detected = None
    if magic:
        try:
            detected = magic.from_file(file_path, mime=True)
        except Exception:
            pass
    if not detected:
        detected, _ = mimetypes.guess_type(file_path)

    if not detected:
        raise ValidationError("Could not determine file type.")

    # Check if detected MIME starts with any allowed type
    matched = any(
        detected.startswith(allowed) or allowed.startswith(detected)
        for allowed in allowed_mime_types
    )
    if not matched:
        raise ValidationError(
            f"Invalid file content. Detected MIME type: {detected}. "
            f"Expected one of: {', '.join(sorted(allowed_mime_types))}"
        )
    return detected


def validate_video(file_obj, max_mb=2048):
    """Full validation pipeline for video uploads."""
    validate_upload_size(file_obj, max_mb)
    ext = validate_extension(file_obj.name, ALLOWED_VIDEO_EXTENSIONS)
    return ext


def validate_image(file_obj, max_mb=100):
    """Full validation pipeline for image uploads."""
    validate_upload_size(file_obj, max_mb)
    ext = validate_extension(file_obj.name, ALLOWED_IMAGE_EXTENSIONS)
    return ext


def validate_audio(file_obj, max_mb=500):
    """Full validation pipeline for audio uploads."""
    validate_upload_size(file_obj, max_mb)
    ext = validate_extension(file_obj.name, ALLOWED_AUDIO_EXTENSIONS)
    return ext


def validate_subtitle(file_obj, max_mb=10):
    """Full validation pipeline for subtitle uploads."""
    validate_upload_size(file_obj, max_mb)
    ext = validate_extension(file_obj.name, ALLOWED_SUBTITLE_EXTENSIONS)
    return ext


def sanitize_filename(name):
    """Sanitize filename to prevent path traversal and injection."""
    import re
    # Remove path separators and null bytes
    name = os.path.basename(name)
    name = name.replace('\x00', '')
    # Allow only safe characters
    name = re.sub(r'[^\w\s\.\-]', '', name)
    # Prevent double extensions / command injection
    name = name.replace(';', '').replace('|', '').replace('&', '').replace('`', '')
    return name


def validate_ffmpeg_command_args(args_list):
    """
    Validate that FFmpeg command arguments don't contain injection attempts.
    Returns cleaned args list.
    """
    dangerous = {';', '|', '&&', '||', '`', '$', '(', ')', '<', '>'}
    cleaned = []
    for arg in args_list:
        if any(d in str(arg) for d in dangerous):
            raise ValidationError(f"Invalid character in argument: {arg}")
        cleaned.append(str(arg))
    return cleaned
