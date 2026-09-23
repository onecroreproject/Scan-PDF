import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_speed_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")

def process_speed(
    video,
    speed=1.0,
    keep_audio=True,
    output_format="mp4",
):
    """
    Adjust video playback speed (slow motion / speed up).
    """
    if not video:
        raise ValueError("Video file is required.")

    try:
        speed = float(speed)
        if speed <= 0:
            speed = 1.0
    except (ValueError, TypeError):
        speed = 1.0

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)

    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_speed_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        speed=speed,
        keep_audio=keep_audio,
        output_format=output_format,
    )

    logger.info("Executing speed change to %sx: %s", speed, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Speed processing failed to create output video.")

    return output_path
