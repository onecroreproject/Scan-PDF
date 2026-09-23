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


def process_split_screen(
    video1,
    video2,
    video3=None,
    video4=None,
    layout="2-side",  # '2-side', '2-stack', '4-grid'
    output_format="mp4",
):
    """
    Combine multiple videos into a synchronized split-screen video.
    Layouts:
      - 2-side: Left / Right side-by-side (scaled to 1920x1080)
      - 2-stack: Top / Bottom stacked (scaled to 1080x1920 or 1920x1080)
      - 4-grid: 2x2 grid
    """
    if not video1 or not video2:
        raise ValueError("At least two video files are required for split screen.")

    v1_path = save_uploaded_file(video1)
    v2_path = save_uploaded_file(video2)
    v3_path = save_uploaded_file(video3) if video3 else None
    v4_path = save_uploaded_file(video4) if video4 else None

    _, outputs_dir, _ = get_video_directories()
    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    inputs = ["-i", str(v1_path), "-i", str(v2_path)]

    if layout == "2-stack":
        # Two videos stacked vertically (each scaled to 1280x360 -> total 1280x720)
        filter_complex = (
            "[0:v]scale=1280:360:force_original_aspect_ratio=increase,crop=1280:360[top];"
            "[1:v]scale=1280:360:force_original_aspect_ratio=increase,crop=1280:360[bot];"
            "[top][bot]vstack=inputs=2[vout];"
            "[0:a][1:a]amix=inputs=2:duration=first[aout]"
        )
    elif layout == "4-grid" and v3_path and v4_path:
        inputs.extend(["-i", str(v3_path), "-i", str(v4_path)])
        filter_complex = (
            "[0:v]scale=640:360:force_original_aspect_ratio=increase,crop=640:360[v0];"
            "[1:v]scale=640:360:force_original_aspect_ratio=increase,crop=640:360[v1];"
            "[2:v]scale=640:360:force_original_aspect_ratio=increase,crop=640:360[v2];"
            "[3:v]scale=640:360:force_original_aspect_ratio=increase,crop=640:360[v3];"
            "[v0][v1]hstack=inputs=2[row0];"
            "[v2][v3]hstack=inputs=2[row1];"
            "[row0][row1]vstack=inputs=2[vout];"
            "[0:a][1:a]amix=inputs=2:duration=first[aout]"
        )
    else:
        # Default: 2-side (each scaled to 640x720 -> total 1280x720)
        filter_complex = (
            "[0:v]scale=640:720:force_original_aspect_ratio=increase,crop=640:720[left];"
            "[1:v]scale=640:720:force_original_aspect_ratio=increase,crop=640:720[right];"
            "[left][right]hstack=inputs=2[vout];"
            "[0:a][1:a]amix=inputs=2:duration=first[aout]"
        )

    command = [
        str(ffmpeg_binary),
        "-y",
        *inputs,
        "-filter_complex", filter_complex,
        "-map", "[vout]",
        "-map", "[aout]",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "21",
        "-pix_fmt", "yuv420p",
        "-c:a", "aac",
        "-b:a", "192k",
        "-shortest",
        "-movflags", "+faststart",
        str(output_path),
    ]

    logger.info("Executing split screen video (layout=%s)", layout)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Split screen generation failed.")

    return output_path
