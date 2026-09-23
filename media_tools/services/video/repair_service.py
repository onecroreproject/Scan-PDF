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


def process_repair(
    video,
    repair_mode="smart",  # 'smart' (try copy then re-encode), 'remux', 'reencode'
    output_format="mp4",
):
    """
    Attempt to repair damaged, truncated, or unplayable video files by rebuilding
    container structures, fixing index tables (MOOV atoms), and ignoring corrupted packets.
    """
    if not video:
        raise ValueError("Video file is required.")

    video_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    # Strategy 1: Remux with ignore_err and faststart (fastest, preserves 100% quality)
    if repair_mode in ("smart", "remux"):
        command_remux = [
            str(ffmpeg_binary),
            "-y",
            "-err_detect", "ignore_err",
            "-i", str(video_path),
            "-c", "copy",
            "-movflags", "+faststart",
            str(output_path),
        ]
        try:
            logger.info("Attempting remux repair: %s", video_path)
            run_ffmpeg(command_remux)
            if output_path.exists() and output_path.stat().st_size > 0:
                return output_path
        except Exception as exc:
            logger.warning("Remux repair failed, attempting deep re-encode repair: %s", exc)

    # Strategy 2: Deep Re-encode (re-indexes, synthesizes timestamps, drops corrupt frames)
    command_reencode = [
        str(ffmpeg_binary),
        "-y",
        "-err_detect", "ignore_err",
        "-fflags", "+genpts+discardcorrupt",
        "-i", str(video_path),
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-c:a", "aac",
        "-b:a", "128k",
        "-movflags", "+faststart",
        str(output_path),
    ]

    logger.info("Executing deep re-encode repair: %s", video_path)
    run_ffmpeg(command_reencode)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("The uploaded video file could not be recovered. The file may be severely corrupted.")

    return output_path
