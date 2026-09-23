import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_loop_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_loop(video, loop_count=2, output_format="mp4"):
    """
    Loop a video N times.
    """
    if not video:
        raise ValueError("Video file is required.")

    try:
        loop_count = max(2, int(loop_count))
    except (ValueError, TypeError):
        loop_count = 2

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_loop_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        loop_count=loop_count,
        output_format=output_format,
    )
    logger.info("Executing loop video (%sx): %s", loop_count, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Loop video processing failed.")

    return output_path
