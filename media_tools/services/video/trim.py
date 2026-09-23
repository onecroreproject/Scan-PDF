import logging
from decimal import Decimal
from pathlib import Path

from django.conf import settings

from media_tools.services.ffmpeg_commands import (
    build_trim_command,
)
from media_tools.services.ffmpeg_runner import (
    run_ffmpeg,
)
from media_tools.services.file_service import (
    create_unique_filename,
    get_video_directories,
    save_uploaded_file,
)
from media_tools.services.media_info import (
    get_video_info,
)


logger = logging.getLogger("media_tools")


ALLOWED_FORMATS = {
    "mp4",
    "mov",
    "webm",
    "avi",
    "mkv",
    "m4v",
    "flv",
    "wmv",
    "gif",
    "mpeg",
}


MAX_VIDEO_SIZE = (
    500
    * 1024
    * 1024
)


MIN_TRIM_DURATION = Decimal(
    "0.001"
)


def process_trim(
    video,
    start_time,
    end_time,
    output_format="mp4",
):
    """
    Extract a selected section from a video.

    Processing flow:

        Upload
          ↓
        Validate
          ↓
        FFprobe
          ↓
        Validate timeline
          ↓
        Create output path
          ↓
        Build FFmpeg command
          ↓
        Run FFmpeg
          ↓
        Verify output
          ↓
        Return output path
    """

    input_path = None
    output_path = None

    try:

        # --------------------------------------------------
        # Basic validation
        # --------------------------------------------------

        if not video:
            raise ValueError(
                "Video file is required."
            )

        if start_time is None:
            raise ValueError(
                "Start time is required."
            )

        if end_time is None:
            raise ValueError(
                "End time is required."
            )

        try:
            start_time = Decimal(
                str(start_time)
            )

            end_time = Decimal(
                str(end_time)
            )

        except (
            ValueError,
            TypeError,
        ) as exc:

            raise ValueError(
                "Invalid trim time."
            ) from exc

        if start_time < 0:
            raise ValueError(
                "Start time cannot be negative."
            )

        if end_time <= start_time:
            raise ValueError(
                "End time must be greater than start time."
            )

        trim_duration = (
            end_time
            - start_time
        )

        if (
            trim_duration
            < MIN_TRIM_DURATION
        ):
            raise ValueError(
                "The selected section is too short."
            )

        # --------------------------------------------------
        # Output format validation
        # --------------------------------------------------

        output_format = (
            str(output_format)
            .lower()
            .strip()
        )

        if output_format not in ALLOWED_FORMATS:
            raise ValueError(
                "Invalid output format."
            )

        # --------------------------------------------------
        # Upload size validation
        # --------------------------------------------------

        video_size = getattr(
            video,
            "size",
            0,
        )

        if video_size <= 0:
            raise ValueError(
                "The uploaded video is empty."
            )

        if video_size > MAX_VIDEO_SIZE:
            raise ValueError(
                "Video file must be smaller than 500 MB."
            )

        logger.info(
            "Starting video trim operation."
        )

        logger.info(
            "Trim settings: "
            "start=%s, end=%s, duration=%s, format=%s",
            start_time,
            end_time,
            trim_duration,
            output_format,
        )

        # --------------------------------------------------
        # Save uploaded video
        # --------------------------------------------------

        input_path = save_uploaded_file(
            video
        )

        if not input_path:
            raise RuntimeError(
                "Unable to save uploaded video."
            )

        input_path = Path(
            input_path
        )

        if not input_path.exists():
            raise RuntimeError(
                "Uploaded video was not saved."
            )

        if input_path.stat().st_size <= 0:
            raise RuntimeError(
                "Uploaded video is empty."
            )

        logger.info(
            "Uploaded video saved: %s",
            input_path,
        )

        # --------------------------------------------------
        # Read source video information
        # --------------------------------------------------

        info = get_video_info(
            input_path
        )

        source_width = int(
            info.get(
                "width",
                0,
            )
        )

        source_height = int(
            info.get(
                "height",
                0,
            )
        )

        source_duration = Decimal(
            str(
                info.get(
                    "duration",
                    0,
                )
                or 0
            )
        )

        if source_width <= 0:
            raise ValueError(
                "Invalid source video width."
            )

        if source_height <= 0:
            raise ValueError(
                "Invalid source video height."
            )

        if source_duration <= 0:
            raise ValueError(
                "Unable to determine video duration."
            )

        logger.info(
            "Source video: %sx%s, duration=%s",
            source_width,
            source_height,
            source_duration,
        )

        # --------------------------------------------------
        # Validate trim against actual duration
        # --------------------------------------------------

        if start_time >= source_duration:
            raise ValueError(
                "Start time is outside the video duration."
            )

        # Clamp end_time instead of hard-rejecting — browser video.duration
        # can differ from FFprobe duration by a few milliseconds due to
        # floating-point representation differences.
        TOLERANCE = Decimal("0.5")  # allow up to 0.5s overshoot
        if end_time > source_duration:
            if end_time - source_duration <= TOLERANCE:
                # Clamp to exact source duration
                end_time = source_duration
            else:
                raise ValueError(
                    f"End time ({float(end_time):.3f}s) cannot be greater than "
                    f"the video duration ({float(source_duration):.3f}s)."
                )

        if end_time <= start_time:
            raise ValueError(
                "End time must be greater than start time."
            )

        actual_duration = (
            end_time
            - start_time
        )

        if actual_duration <= 0:
            raise ValueError(
                "Selected video duration must be positive."
            )

        # --------------------------------------------------
        # Get output directory
        # --------------------------------------------------

        _, outputs_dir, _ = (
            get_video_directories()
        )

        outputs_dir = Path(
            outputs_dir
        )

        if not outputs_dir.exists():
            raise RuntimeError(
                "Output directory does not exist."
            )

        # --------------------------------------------------
        # Create output filename
        # --------------------------------------------------

        extension = (
            f".{output_format}"
        )

        output_path = (
            outputs_dir
            / create_unique_filename(
                extension
            )
        )

        output_path = Path(
            output_path
        )

        # --------------------------------------------------
        # FFmpeg binary
        # --------------------------------------------------

        ffmpeg_binary = getattr(
            settings,
            "FFMPEG_BINARY",
            "ffmpeg",
        )

        if not ffmpeg_binary:
            raise RuntimeError(
                "FFmpeg binary is not configured."
            )

        # --------------------------------------------------
        # Build command
        # --------------------------------------------------

        command = build_trim_command(
            ffmpeg_binary,
            input_path,
            output_path,
            start_time,
            end_time,
            output_format,
        )

        if not command:
            raise RuntimeError(
                "Unable to build FFmpeg command."
            )

        logger.info(
            "Running FFmpeg trim operation."
        )

        # --------------------------------------------------
        # Run FFmpeg
        # --------------------------------------------------

        run_ffmpeg(
            command
        )

        # --------------------------------------------------
        # Verify output
        # --------------------------------------------------

        if not output_path.exists():
            raise RuntimeError(
                "FFmpeg did not create output."
            )

        output_size = (
            output_path.stat().st_size
        )

        if output_size <= 0:
            raise RuntimeError(
                "Generated video is empty."
            )

        logger.info(
            "Trim completed successfully: %s",
            output_path,
        )

        return output_path

    except ValueError:
        logger.warning(
            "Trim validation failed."
        )
        raise

    except Exception as exc:

        logger.exception(
            "Trim processing failed."
        )

        raise RuntimeError(
            "Unable to trim the video."
        ) from exc
