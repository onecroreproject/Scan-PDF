import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_commands import (
    build_remove_audio_command,
    build_extract_audio_command,
    build_volume_command,
)
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def process_remove_audio(video, output_format="mp4"):
    """
    Remove all audio tracks from a video (mute).
    """
    if not video:
        raise ValueError("Video file is required.")

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_remove_audio_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        output_format=output_format,
    )
    logger.info("Executing remove audio: %s", input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Remove audio failed to create output file.")

    return output_path


def process_extract_audio(video, audio_format="mp3", bitrate="192k"):
    """
    Extract audio track from video to MP3, WAV, AAC, M4A, OGG.
    """
    if not video:
        raise ValueError("Video file is required.")

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    audio_format = audio_format.lower().lstrip(".")
    extension = f".{audio_format}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_extract_audio_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        audio_format=audio_format,
        bitrate=bitrate,
    )
    logger.info("Executing extract audio (%s): %s", audio_format, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Extract audio failed to create output file.")

    return output_path


def process_change_volume(video, volume=100, output_format="mp4"):
    """
    Adjust audio volume of video (percentage: 0 to 300%).
    """
    if not video:
        raise ValueError("Video file is required.")

    try:
        vol_pct = float(volume)
        volume_ratio = max(0.0, vol_pct / 100.0)
    except (ValueError, TypeError):
        volume_ratio = 1.0

    input_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    command = build_volume_command(
        ffmpeg_binary=ffmpeg_binary,
        input_path=input_path,
        output_path=output_path,
        volume_ratio=volume_ratio,
        output_format=output_format,
    )
    logger.info("Executing volume adjust (%s%%): %s", volume, input_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Volume adjustment failed to create output video.")

    return output_path
