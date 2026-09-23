import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import build_fps_command
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")

def process_fps(
    video,
    fps=30,
    output_format="mp4",
):
    """
    Change video frame rate (FPS).
    """
    if not video:
        raise ValueError("Video file is required.")

    try:
        fps = int(fps)
        if fps <= 0:
            fps = 30
    except (ValueError, TypeError):
        fps = 30

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)

    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_fps_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        fps=fps,
        output_format=output_format,
    )

    logger.info("Executing FPS change to %s fps: %s", fps, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("FPS change failed to create output video.")

    return output_path
