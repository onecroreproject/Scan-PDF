import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import (
    build_video_to_gif_command,
    build_gif_to_video_command,
)
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_video_to_gif(video, fps=15, width=480):
    """
    Convert video to an animated GIF using high quality palette.
    """
    if not video:
        raise ValueError("Video file is required.")

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    output_path = outputs_dir / create_unique_filename(".gif")
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_video_to_gif_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        fps=fps,
        width=width,
    )
    logger.info("Executing video to GIF: %s", input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Video to GIF conversion failed.")

    return output_path


def process_gif_to_video(video, output_format="mp4"):
    """
    Convert animated GIF to smooth playable MP4 video.
    """
    if not video:
        raise ValueError("GIF file is required.")

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_gif_to_video_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        output_format=output_format,
    )
    logger.info("Executing GIF to video: %s", input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("GIF to video conversion failed.")

    return output_path
