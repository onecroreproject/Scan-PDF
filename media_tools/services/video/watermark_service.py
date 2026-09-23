import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_watermark_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_watermark(
    video,
    watermark_image,
    position="bottom-right",
    opacity=0.8,
    scale_percent=15,
    output_format="mp4",
):
    """
    Overlay image watermark on video.
    """
    if not video:
        raise ValueError("Video file is required.")
    if not watermark_image:
        raise ValueError("Watermark image is required.")

    input_path = save_uploaded_file(video)
    watermark_path = save_uploaded_file(watermark_image)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_watermark_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        watermark_path=watermark_path,
        output_path=output_path,
        position=position,
        opacity=opacity,
        scale_percent=scale_percent,
        output_format=output_format,
    )
    logger.info("Executing watermark overlay: %s", input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Watermark processing failed.")

    return output_path
