import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_compress_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")

def process_compress(
    video,
    compression_level="medium",
    output_format="mp4",
):
    """
    Compress video file size.
    """
    if not video:
        raise ValueError("Video file is required.")

    if compression_level not in ("low", "medium", "high"):
        compression_level = "medium"

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)

    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_compress_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        compression_level=compression_level,
        output_format=output_format,
    )

    logger.info("Executing video compression (%s): %s", compression_level, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Video compression failed to create output video.")

    return output_path
