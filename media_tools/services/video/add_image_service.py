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


def process_add_image(
    video,
    image,
    position="bottom-right",
    opacity=0.85,
    scale_percent=20,
    x_percent=80,
    y_percent=80,
    start_time=None,
    end_time=None,
    output_format="mp4",
):
    """
    Overlay image/logo onto video with position, scale, opacity, and timing controls.
    """
    if not video:
        raise ValueError("Video file is required.")
    if not image:
        raise ValueError("Image file is required.")

    video_path = save_uploaded_file(video)
    image_path = save_uploaded_file(image)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    try:
        op = max(0.05, min(1.0, float(opacity)))
    except (ValueError, TypeError):
        op = 0.85

    try:
        sc = max(5, min(100, float(scale_percent))) / 100.0
    except (ValueError, TypeError):
        sc = 0.20

    pos_map = {
        "top-left": "25:25",
        "top-center": "(main_w-overlay_w)/2:25",
        "top-right": "main_w-overlay_w-25:25",
        "center": "(main_w-overlay_w)/2:(main_h-overlay_h)/2",
        "bottom-left": "25:main_h-overlay_h-25",
        "bottom-center": "(main_w-overlay_w)/2:main_h-overlay_h-25",
        "bottom-right": "main_w-overlay_w-25:main_h-overlay_h-25",
    }

    if position in pos_map:
        overlay_pos = pos_map[position]
    else:
        try:
            xp = max(0, min(100, float(x_percent))) / 100.0
            yp = max(0, min(100, float(y_percent))) / 100.0
            overlay_pos = f"(main_w*{xp:.2f})-(overlay_w/2):(main_h*{yp:.2f})-(overlay_h/2)"
        except (ValueError, TypeError):
            overlay_pos = "main_w-overlay_w-25:main_h-overlay_h-25"

    enable_clause = ""
    if start_time is not None and str(start_time).strip() != "":
        try:
            st = float(start_time)
            if end_time is not None and str(end_time).strip() != "":
                et = float(end_time)
                enable_clause = f":enable='between(t,{st:.2f},{et:.2f})'"
            else:
                enable_clause = f":enable='gte(t,{st:.2f})'"
        except (ValueError, TypeError):
            pass

    filter_complex = (
        f"[1:v]scale=iw*{sc:.2f}:-1,format=rgba,colorchannelmixer=aa={op:.2f}[img];"
        f"[0:v][img]overlay={overlay_pos}{enable_clause}[vout]"
    )

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(video_path),
        "-i", str(image_path),
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

    logger.info("Executing add image overlay: %s", video_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Add image overlay failed to generate output video.")

    return output_path
