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


def process_add_audio(
    video,
    audio,
    mode="replace",  # 'replace' or 'mix'
    audio_volume=1.0,
    video_volume=1.0,
    output_format="mp4",
):
    """
    Add audio to video: replace original audio or mix with original audio.
    """
    if not video:
        raise ValueError("Video file is required.")
    if not audio:
        raise ValueError("Audio file is required.")

    video_path = save_uploaded_file(video)
    audio_path = save_uploaded_file(audio)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    try:
        a_vol = float(audio_volume)
    except (ValueError, TypeError):
        a_vol = 1.0

    try:
        v_vol = float(video_volume)
    except (ValueError, TypeError):
        v_vol = 1.0

    if mode == "mix":
        # Mix original video audio with new audio track
        filter_complex = (
            f"[0:a]volume={v_vol}[a0];"
            f"[1:a]volume={a_vol}[a1];"
            f"[a0][a1]amix=inputs=2:duration=first:dropout_transition=2[aout]"
        )
        command = [
            str(ffmpeg_binary),
            "-y",
            "-i", str(video_path),
            "-i", str(audio_path),
            "-filter_complex", filter_complex,
            "-map", "0:v",
            "-map", "[aout]",
            "-c:v", "copy",
            "-c:a", "aac",
            "-b:a", "192k",
            "-shortest",
            "-movflags", "+faststart",
            str(output_path),
        ]
    else:
        # Replace original audio completely with new audio
        if a_vol != 1.0:
            filter_complex = f"[1:a]volume={a_vol}[aout]"
            command = [
                str(ffmpeg_binary),
                "-y",
                "-i", str(video_path),
                "-i", str(audio_path),
                "-filter_complex", filter_complex,
                "-map", "0:v",
                "-map", "[aout]",
                "-c:v", "copy",
                "-c:a", "aac",
                "-b:a", "192k",
                "-shortest",
                "-movflags", "+faststart",
                str(output_path),
            ]
        else:
            command = [
                str(ffmpeg_binary),
                "-y",
                "-i", str(video_path),
                "-i", str(audio_path),
                "-map", "0:v",
                "-map", "1:a",
                "-c:v", "copy",
                "-c:a", "aac",
                "-b:a", "192k",
                "-shortest",
                "-movflags", "+faststart",
                str(output_path),
            ]

    logger.info("Executing add audio (mode=%s): %s + %s", mode, video_path, audio_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Add audio failed to generate output video.")

    return output_path
