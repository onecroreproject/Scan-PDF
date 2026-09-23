import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_convert_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")

def process_convert(
    video,
    target_format="mp4",
):
    """
    Convert video to a different container/format.
    """
    if not video:
        raise ValueError("Video file is required.")

    target_format = str(target_format).lower().lstrip(".")
    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{target_format}"
    output_path = outputs_dir / create_unique_filename(extension)

    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_convert_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        target_format=target_format,
    )

    logger.info("Executing format conversion to %s: %s", target_format, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError(f"Conversion to {target_format} failed to create output file.")

    return output_path
