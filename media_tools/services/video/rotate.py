import logging

from django.conf import settings

from media_tools.services.ffmpeg_commands import (
    build_rotate_command,
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
    "gif",
    "m4v",
}


ALLOWED_ROTATIONS = {
    "90",
    "180",
    "270",
}


def process_rotate(
    video,
    rotation="90",
    output_format="mp4",
):
    """
    Rotate a video using FFmpeg.

    rotation:
        90  = clockwise 90 degrees
        180 = upside down
        270 = clockwise 270 degrees
              / counter-clockwise 90 degrees

    Returns:
        pathlib.Path: Generated output video path.
    """

    input_path = None
    output_path = None

    try:
        # -------------------------------------------------
        # Validate video
        # -------------------------------------------------

        if not video:
            raise ValueError(
                "Video file is required."
            )

        # -------------------------------------------------
        # Validate rotation
        # -------------------------------------------------

        rotation = str(rotation).strip()

        if rotation not in ALLOWED_ROTATIONS:
            raise ValueError(
                "Invalid rotation angle."
            )

        # -------------------------------------------------
        # Validate output format
        # -------------------------------------------------

        output_format = (
            str(output_format)
            .strip()
            .lower()
        )

        if output_format not in ALLOWED_FORMATS:
            raise ValueError(
                "Invalid output format."
            )

        logger.info(
            "Starting rotate operation."
        )

        logger.info(
            "Rotation settings: rotation=%s, format=%s",
            rotation,
            output_format,
        )

        # -------------------------------------------------
        # Save uploaded video
        # -------------------------------------------------

        input_path = save_uploaded_file(
            video
        )

        # -------------------------------------------------
        # Read source metadata
        # -------------------------------------------------

        info = get_video_info(
            input_path
        )

        source_width = int(
            info.get("width", 0)
        )

        source_height = int(
            info.get("height", 0)
        )

        if source_width <= 0:
            raise ValueError(
                "Invalid source video width."
            )

        if source_height <= 0:
            raise ValueError(
                "Invalid source video height."
            )

        logger.info(
            "Source dimensions: %sx%s",
            source_width,
            source_height,
        )

        # -------------------------------------------------
        # Get output directories
        # -------------------------------------------------

        _, outputs_dir, _ = (
            get_video_directories()
        )

        # -------------------------------------------------
        # Create output filename
        # -------------------------------------------------

        extension = (
            f".{output_format}"
        )

        output_path = (
            outputs_dir
            / create_unique_filename(
                extension
            )
        )

        # -------------------------------------------------
        # FFmpeg binary
        # -------------------------------------------------

        ffmpeg_binary = getattr(
            settings,
            "FFMPEG_BINARY",
            "ffmpeg",
        )

        # -------------------------------------------------
        # Build FFmpeg command
        # -------------------------------------------------

        command = build_rotate_command(
            ffmpeg_binary,
            input_path,
            output_path,
            rotation,
            output_format,
        )

        logger.info(
            "Running FFmpeg rotate command."
        )

        # -------------------------------------------------
        # Run FFmpeg
        # -------------------------------------------------

        run_ffmpeg(
            command
        )

        # -------------------------------------------------
        # Verify output
        # -------------------------------------------------

        if not output_path.exists():
            raise RuntimeError(
                "FFmpeg did not create output."
            )

        if output_path.stat().st_size <= 0:
            raise RuntimeError(
                "Generated video is empty."
            )

        logger.info(
            "Rotate completed successfully: %s",
            output_path,
        )

        return output_path

    except ValueError:
        logger.warning(
            "Rotate validation failed."
        )
        raise

    except Exception as exc:
        logger.exception(
            "Rotate processing failed."
        )

        raise RuntimeError(
            "Unable to rotate the video."
        ) from exc
