import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_fade_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.media_info import get_video_info
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_fade(
    video,
    fade_in_duration=1.0,
    fade_out_duration=1.0,
    color="black",
    output_format="mp4",
):
    """
    Apply fade in and/or fade out transitions.
    """
    if not video:
        raise ValueError("Video file is required.")

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    # Get video duration from media_info
    total_duration = 10.0
    try:
        info = get_video_info(input_path)
        if info and "duration" in info:
            total_duration = float(info["duration"])
    except Exception as exc:
        logger.warning("Could not probe video duration for fade: %s", exc)

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_fade_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        fade_in_duration=fade_in_duration,
        fade_out_duration=fade_out_duration,
        total_duration=total_duration,
        color=color,
        output_format=output_format,
    )
    logger.info("Executing video fade (in=%s, out=%s): %s", fade_in_duration, fade_out_duration, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Video fade processing failed.")

    return output_path
