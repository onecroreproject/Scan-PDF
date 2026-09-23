import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_color_filter_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_effects(
    video,
    brightness=0.0,
    contrast=1.0,
    saturation=1.0,
    filter_type="none",
    output_format="mp4",
):
    """
    Apply brightness, contrast, saturation, or artistic filters.
    """
    if not video:
        raise ValueError("Video file is required.")

    try:
        brightness = float(brightness)
        contrast = float(contrast)
        saturation = float(saturation)
    except (ValueError, TypeError):
        brightness, contrast, saturation = 0.0, 1.0, 1.0

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_color_filter_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        brightness=brightness,
        contrast=contrast,
        saturation=saturation,
        filter_type=filter_type,
        output_format=output_format,
    )
    logger.info("Executing video effects (%s, b=%s, c=%s, s=%s): %s", filter_type, brightness, contrast, saturation, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Video effects processing failed.")

    return output_path
