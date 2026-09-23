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


ASPECT_DIMENSIONS = {
    "16:9": (1920, 1080),
    "9:16": (1080, 1920),
    "1:1": (1080, 1080),
    "4:5": (1080, 1350),
    "21:9": (2560, 1080),
}


def process_background(
    video,
    bg_type="blur",  # 'blur', 'color'
    bg_color="#000000",
    aspect_ratio="9:16",
    output_format="mp4",
):
    """
    Apply modern background effects to video, such as blurred background
    or stylized solid color padding for social media aspect ratios (9:16 TikTok/Reels, 1:1, etc.).
    """
    if not video:
        raise ValueError("Video file is required.")

    video_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    target_w, target_h = ASPECT_DIMENSIONS.get(aspect_ratio, (1080, 1920))

    if bg_type == "blur":
        # Background is the video scaled to fill and heavily blurred; foreground is original video fitted inside
        filter_complex = (
            f"[0:v]scale={target_w}:{target_h}:force_original_aspect_ratio=increase,"
            f"crop={target_w}:{target_h},boxblur=25:5[bg];"
            f"[0:v]scale={target_w}:{target_h}:force_original_aspect_ratio=decrease[fg];"
            f"[bg][fg]overlay=(W-w)/2:(H-h)/2[vout]"
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
            "-crf", "21",
            "-pix_fmt", "yuv420p",
            "-c:a", "copy",
            "-movflags", "+faststart",
            str(output_path),
        ]
    else:
        # Solid color padding
        color_hex = str(bg_color).strip().lstrip("#")
        if len(color_hex) == 6:
            color_param = f"0x{color_hex}"
        else:
            color_param = "black"

        vf = (
            f"scale={target_w}:{target_h}:force_original_aspect_ratio=decrease,"
            f"pad={target_w}:{target_h}:(ow-iw)/2:(oh-ih)/2:color={color_param}"
        )
        command = [
            str(ffmpeg_binary),
            "-y",
            "-i", str(video_path),
            "-vf", vf,
            "-c:v", "libx264",
            "-preset", "medium",
            "-crf", "21",
            "-pix_fmt", "yuv420p",
            "-c:a", "copy",
            "-movflags", "+faststart",
            str(output_path),
        ]

    logger.info("Executing background tool (%s, ratio=%s): %s", bg_type, aspect_ratio, video_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Background processing failed to generate output video.")

    return output_path
