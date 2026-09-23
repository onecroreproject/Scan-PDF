import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_frame_extract_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_extract_frame(video, time_offset=0.0, image_format="jpg"):
    """
    Extract a high quality still frame / screenshot from video.
    """
    if not video:
        raise ValueError("Video file is required.")

    try:
        time_offset = max(0.0, float(time_offset))
    except (ValueError, TypeError):
        time_offset = 0.0

    image_format = str(image_format).lower().lstrip(".")
    if image_format not in ("jpg", "jpeg", "png", "webp"):
        image_format = "jpg"

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{image_format}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_frame_extract_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        time_offset=time_offset,
    )
    logger.info("Executing frame extract at %ss: %s", time_offset, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Frame extraction failed.")

    return output_path
