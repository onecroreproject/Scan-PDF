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


def process_stabilize(
    video,
    strength="medium",  # 'mild', 'medium', 'strong'
    output_format="mp4",
):
    """
    Stabilize shaky video using the FFmpeg deshake filter.
    """
    if not video:
        raise ValueError("Video file is required.")

    video_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    strength_map = {
        "mild": "rx=16:ry=16:blocksize=32:edge=mirror",
        "medium": "rx=32:ry=32:blocksize=32:edge=mirror",
        "strong": "rx=64:ry=64:blocksize=64:edge=mirror",
    }
    deshake_params = strength_map.get(strength.lower(), strength_map["medium"])

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(video_path),
        "-vf", f"deshake={deshake_params}",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "20",
        "-pix_fmt", "yuv420p",
        "-c:a", "copy",
        "-movflags", "+faststart",
        str(output_path),
    ]

    logger.info("Executing stabilize video (strength=%s): %s", strength, video_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Video stabilization failed to generate output file.")

    return output_path
