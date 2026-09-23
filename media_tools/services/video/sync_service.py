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


def process_sync(
    video,
    audio_offset=0.0,  # Offset in seconds (negative: audio ahead, positive: audio delayed)
    output_format="mp4",
):
    """
    Fix audio/video sync issues by shifting the audio track relative to video.
    Positive offset = delay audio (plays later).
    Negative offset = advance audio (plays earlier).
    """
    if not video:
        raise ValueError("Video file is required.")

    video_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    try:
        offset_sec = float(audio_offset)
    except (ValueError, TypeError):
        offset_sec = 0.0

    offset_ms = int(round(offset_sec * 1000))

    if offset_ms > 0:
        # Delay audio using adelay
        af_filter = f"adelay={offset_ms}|{offset_ms}"
    elif offset_ms < 0:
        # Advance audio by trimming start
        trim_sec = abs(offset_sec)
        af_filter = f"atrim=start={trim_sec:.3f},asetpts=PTS-STARTPTS"
    else:
        af_filter = "anull"

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(video_path),
        "-af", af_filter,
        "-c:v", "copy",
        "-c:a", "aac",
        "-b:a", "192k",
        "-movflags", "+faststart",
        str(output_path),
    ]

    logger.info("Executing audio/video sync (offset=%sms): %s", offset_ms, video_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Audio/video synchronization failed to generate output file.")

    return output_path
