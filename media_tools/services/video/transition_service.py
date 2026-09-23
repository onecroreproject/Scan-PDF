import logging
from pathlib import Path
from django.conf import settings
from media_tools.services.ffmpeg_runner import run_ffmpeg
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)
from media_tools.services.media_info import get_video_info

logger = logging.getLogger("media_tools")


def process_transition(
    video,
    transition_type="fade-both",  # 'fade-in', 'fade-out', 'fade-both', 'fade-white'
    duration=1.0,
    color="black",
    output_format="mp4",
):
    """
    Apply smooth video transitions (fade in, fade out, or both) with customizable duration and color.
    """
    if not video:
        raise ValueError("Video file is required.")

    video_path = save_uploaded_file(video)
    _, outputs_dir, _ = get_video_directories()

    extension = f".{output_format.lower().lstrip('.')}"
    output_path = outputs_dir / create_unique_filename(extension)
    ffmpeg_binary = getattr(settings, "FFMPEG_BINARY", "ffmpeg")

    try:
        dur = max(0.2, min(5.0, float(duration)))
    except (ValueError, TypeError):
        dur = 1.0

    # Get video duration
    try:
        info = get_video_info(video_path)
        total_duration = float(info.get("duration", 10.0))
    except Exception:
        total_duration = 10.0

    fade_color = "white" if transition_type == "fade-white" else color

    vf_parts = []
    af_parts = []

    if transition_type in ("fade-in", "fade-both", "fade-white"):
        vf_parts.append(f"fade=t=in:st=0:d={dur:.2f}:c={fade_color}")
        af_parts.append(f"afade=t=in:st=0:d={dur:.2f}")

    if transition_type in ("fade-out", "fade-both", "fade-white"):
        out_start = max(0.0, total_duration - dur)
        vf_parts.append(f"fade=t=out:st={out_start:.2f}:d={dur:.2f}:c={fade_color}")
        af_parts.append(f"afade=t=out:st={out_start:.2f}:d={dur:.2f}")

    if not vf_parts:
        vf_parts.append(f"fade=t=in:st=0:d={dur:.2f}:c=black")

    vf_str = ",".join(vf_parts)
    af_str = ",".join(af_parts) if af_parts else "anull"

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(video_path),
        "-vf", vf_str,
        "-af", af_str,
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "20",
        "-pix_fmt", "yuv420p",
        "-c:a", "aac",
        "-b:a", "192k",
        "-movflags", "+faststart",
        str(output_path),
    ]

    logger.info("Executing transition (%s, dur=%.2fs): %s", transition_type, dur, video_path)
    run_ffmpeg(command)

    if not output_path.exists() or output_path.stat().st_size <= 0:
        raise RuntimeError("Video transition processing failed to generate output file.")

    return output_path
