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


def process_blur(
    video,
    blur_type="full",
    intensity=15,
    x=0,
    y=0,
    width=0,
    height=0,
    output_format="mp4",
):
    """
    Blur entire video or a selected rectangular region.
    """
    if not video:
        raise ValueError("Video file is required.")

    video_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    try:
        rad = max(1, min(60, int(intensity)))
    except (ValueError, TypeError):
        rad = 15

    if blur_type == "region":
        try:
            rx = max(0, int(x))
            ry = max(0, int(y))
            rw = max(10, int(width))
            rh = max(10, int(height))
        except (ValueError, TypeError):
            rx, ry, rw, rh = 0, 0, 100, 100

        # Region blur: crop the area, blur it, and overlay it back on top of original video
        filter_complex = (
            f"[0:v]crop={rw}:{rh}:{rx}:{ry},boxblur={rad}:1[blurred];"
            f"[0:v][blurred]overlay={rx}:{ry}[vout]"
        )
        command = [
            str(ffmpeg_binary),
            "-y",
            "-i", str(video_path),
            "-filter_complex", filter_complex,
            "-map", "[vout]",
            "-map", "0:a?",
            "-c:v", "libx264",
            "-preset", "medium",
            "-crf", "20",
            "-pix_fmt", "yuv420p",
            "-c:a", "copy",
            "-movflags", "+faststart",
            str(output_path),
        ]
    else:
        # Full video blur
        vf_filter = f"boxblur={rad}:1"
        command = [
            str(ffmpeg_binary),
            "-y",
            "-i", str(video_path),
            "-vf", vf_filter,
            "-c:v", "libx264",
            "-preset", "medium",
            "-crf", "20",
            "-pix_fmt", "yuv420p",
            "-c:a", "copy",
            "-movflags", "+faststart",
            str(output_path),
        ]

    logger.info("Executing blur video (%s, rad=%s): %s", blur_type, rad, video_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Blur video failed to generate output file.")

    return output_path
