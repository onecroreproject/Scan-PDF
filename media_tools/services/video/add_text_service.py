import logging
import re
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)

logger = logging.getLogger("media_tools")


def _escape_ffmpeg_text(text: str) -> str:
    """
    Escape special characters for FFmpeg drawtext filter:
    colon (:), backslash (\), single quote ('), and percent (%).
    """
    if not text:
        return ""
    # Escape backslash first
    text = text.replace("\\", "\\\\")
    # Escape single quotes and colons
    text = text.replace("'", "'\\''").replace(":", "\\:")
    text = text.replace("%", "\\%")
    return text


def process_add_text(
    video,
    text="Sample Text",
    font_size=36,
    font_color="#ffffff",
    position="bottom-center",
    bg_box=False,
    x_percent=50,
    y_percent=85,
    start_time=None,
    end_time=None,
    output_format="mp4",
):
    """
    Overlay text onto video using FFmpeg drawtext filter.
    Supports preset positions or interactive percentage coordinates.
    """
    if not video:
        raise ValueError("Video file is required.")
    if not text or not str(text).strip():
        raise ValueError("Text overlay cannot be empty.")

    video_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    clean_text = _escape_ffmpeg_text(str(text).strip())

    try:
        f_size = max(12, min(140, int(font_size)))
    except (ValueError, TypeError):
        f_size = 36

    # Normalize font color
    color = str(font_color).strip()
    if color.startswith("#"):
        color = f"0x{color[1:]}"

    # Calculate position coordinates
    pos_map = {
        "top-left": ("30", "30"),
        "top-center": ("(w-text_w)/2", "30"),
        "top-right": ("w-text_w-30", "30"),
        "center": ("(w-text_w)/2", "(h-text_h)/2"),
        "bottom-left": ("30", "h-text_h-30"),
        "bottom-center": ("(w-text_w)/2", "h-text_h-40"),
        "bottom-right": ("w-text_w-30", "h-text_h-30"),
    }

    if position in pos_map:
        x_expr, y_expr = pos_map[position]
    else:
        # Interactive percentage coordinates from frontend editor
        try:
            xp = max(0, min(100, float(x_percent))) / 100.0
            yp = max(0, min(100, float(y_percent))) / 100.0
            x_expr = f"(w*{xp:.2f})-(text_w/2)"
            y_expr = f"(h*{yp:.2f})-(text_h/2)"
        except (ValueError, TypeError):
            x_expr, y_expr = ("(w-text_w)/2", "h-text_h-40")

    drawtext_parts = [
        f"text='{clean_text}'",
        f"fontsize={f_size}",
        f"fontcolor={color}",
        f"x={x_expr}",
        f"y={y_expr}",
    ]

    if bg_box:
        drawtext_parts.append("box=1:boxcolor=black@0.6:boxborderw=8")

    # Time limits
    if start_time is not None and str(start_time).strip() != "":
        try:
            st = float(start_time)
            if end_time is not None and str(end_time).strip() != "":
                et = float(end_time)
                drawtext_parts.append(f"enable='between(t,{st:.2f},{et:.2f})'")
            else:
                drawtext_parts.append(f"enable='gte(t,{st:.2f})'")
        except (ValueError, TypeError):
            pass

    filter_str = "drawtext=" + ":".join(drawtext_parts)

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(video_path),
        "-vf", filter_str,
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "20",
        "-pix_fmt", "yuv420p",
        "-c:a", "copy",
        "-movflags", "+faststart",
        str(output_path),
    ]

    logger.info("Executing add text overlay: %s", video_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Add text failed to generate output video.")

    return output_path
