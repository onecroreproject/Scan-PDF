import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_reverse_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_reverse(video, reverse_audio=True, output_format="mp4"):
    """
    Reverse video and audio playback.
    """
    if not video:
        raise ValueError("Video file is required.")

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_reverse_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        reverse_audio=reverse_audio,
        output_format=output_format,
    )
    logger.info("Executing reverse video: %s", input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Reverse video processing failed.")

    return output_path
