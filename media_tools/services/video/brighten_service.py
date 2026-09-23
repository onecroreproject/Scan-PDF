import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_brighten(
    video,
    brightness=0.15,
    contrast=1.0,
    saturation=1.0,
    gamma=1.0,
    output_format="mp4",
):
    """
    Adjust video brightness, contrast, and saturation using the FFmpeg eq filter.
    """
    if not video:
        raise ValueError("Video file is required.")

    video_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    try:
        b = max(-1.0, min(1.0, float(brightness)))
    except (ValueError, TypeError):
        b = 0.15

    try:
        c = max(0.1, min(3.0, float(contrast)))
    except (ValueError, TypeError):
        c = 1.0

    try:
        s = max(0.0, min(3.0, float(saturation)))
    except (ValueError, TypeError):
        s = 1.0

    try:
        g = max(0.1, min(3.0, float(gamma)))
    except (ValueError, TypeError):
        g = 1.0

    eq_filter = f"eq=brightness={b:.2f}:contrast={c:.2f}:saturation={s:.2f}:gamma={g:.2f}"

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(video_path),
        "-vf", eq_filter,
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "20",
        "-pix_fmt", "yuv420p",
        "-c:a", "copy",
        "-movflags", "+faststart",
        str(output_path),
    ]

    logger.info("Executing brighten video: %s (b=%.2f, c=%.2f)", video_path, b, c)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Brighten video failed to generate output file.")

    return output_path
