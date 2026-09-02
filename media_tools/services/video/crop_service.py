import logging
from pathlib import Path

from django.conf import settings

from media_tools.services.ffmpeg_commands import (
    build_crop_command,
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


def _to_int(value, name):
    """
    Convert crop value to integer safely.
    """

    try:
        return int(round(float(value)))

    except (
        TypeError,
        ValueError,
    ) as exc:

        raise ValueError(
            f"Invalid {name}."
        ) from exc


def _make_even(value):
    """
    Make crop dimension even.

    This improves compatibility with
    H.264/YUV420 output.
    """

    value = int(value)

    if value < 2:
        return 2

    return value - (
        value % 2
    )


def _validate_crop(
    x,
    y,
    width,
    height,
    video_width,
    video_height,
):
    """
    Validate crop against the actual
    source video dimensions.
    """

    x = _to_int(
        x,
        "crop X",
    )

    y = _to_int(
        y,
        "crop Y",
    )

    width = _to_int(
        width,
        "crop width",
    )

    height = _to_int(
        height,
        "crop height",
    )

    logger.info(
        "Requested crop: "
        "x=%s y=%s width=%s height=%s",
        x,
        y,
        width,
        height,
    )

    # -----------------------------------------
    # Basic validation
    # -----------------------------------------

    if x < 0:
        raise ValueError(
            "Crop X cannot be negative."
        )

    if y < 0:
        raise ValueError(
            "Crop Y cannot be negative."
        )

    if width <= 0:
        raise ValueError(
            "Crop width must be positive."
        )

    if height <= 0:
        raise ValueError(
            "Crop height must be positive."
        )

    # -----------------------------------------
    # Source video validation
    # -----------------------------------------

    if video_width <= 0:
        raise ValueError(
            "Invalid video width."
        )

    if video_height <= 0:
        raise ValueError(
            "Invalid video height."
        )

    # -----------------------------------------
    # Crop position
    # -----------------------------------------

    if x >= video_width:
        raise ValueError(
            "Crop X is outside the video."
        )

    if y >= video_height:
        raise ValueError(
            "Crop Y is outside the video."
        )

    # -----------------------------------------
    # Crop dimensions
    # -----------------------------------------

    if x + width > video_width:

        width = (
            video_width - x
        )

    if y + height > video_height:

        height = (
            video_height - y
        )

    # -----------------------------------------
    # Even dimensions
    # -----------------------------------------

    width = _make_even(
        width
    )

    height = _make_even(
        height
    )

    # -----------------------------------------
    # Final boundary check
    # -----------------------------------------

    if x + width > video_width:

        width -= 2

    if y + height > video_height:

        height -= 2

    if width < 2:
        raise ValueError(
            "Crop width is too small."
        )

    if height < 2:
        raise ValueError(
            "Crop height is too small."
        )

    logger.info(
        "Validated crop: "
        "x=%s y=%s width=%s height=%s",
        x,
        y,
        width,
        height,
    )

    return (
        x,
        y,
        width,
        height,
    )


def process_crop(
    video,
    x,
    y,
    width,
    height,
):
    """
    Process video crop using FFmpeg.

    JavaScript provides the crop area.
    Backend validates it again.
    """

    input_path = None
    output_path = None

    try:

        # -----------------------------------------
        # Upload validation
        # -----------------------------------------

        if not video:
            raise ValueError(
                "Video file is required."
            )

        logger.info(
            "Starting crop operation."
        )

        # -----------------------------------------
        # Convert values
        # -----------------------------------------

        x = _to_int(
            x,
            "crop X",
        )

        y = _to_int(
            y,
            "crop Y",
        )

        width = _to_int(
            width,
            "crop width",
        )

        height = _to_int(
            height,
            "crop height",
        )

        # -----------------------------------------
        # Save uploaded file
        # -----------------------------------------

        input_path = save_uploaded_file(
            video
        )

        input_path = Path(
            input_path
        )

        logger.info(
            "Input video: %s",
            input_path,
        )

        # -----------------------------------------
        # File check
        # -----------------------------------------

        if not input_path.exists():
            raise RuntimeError(
                "Uploaded video was not saved."
            )

        if input_path.stat().st_size <= 0:
            raise RuntimeError(
                "Uploaded video is empty."
            )

        # -----------------------------------------
        # Get actual video information
        # -----------------------------------------

        info = get_video_info(
            input_path
        )

        video_width = int(
            info["width"]
        )

        video_height = int(
            info["height"]
        )

        logger.info(
            "Source dimensions: %sx%s",
            video_width,
            video_height,
        )

        logger.info(
            "Video codec: %s",
            info.get(
                "video_codec"
            ),
        )

        logger.info(
            "Audio codec: %s",
            info.get(
                "audio_codec"
            ),
        )

        # -----------------------------------------
        # Validate crop
        # -----------------------------------------

        (
            x,
            y,
            width,
            height,
        ) = _validate_crop(
            x=x,
            y=y,
            width=width,
            height=height,
            video_width=video_width,
            video_height=video_height,
        )

        # -----------------------------------------
        # Output directory
        # -----------------------------------------

        _, outputs_dir, _ = (
            get_video_directories()
        )

        outputs_dir = Path(
            outputs_dir
        )

        outputs_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        # -----------------------------------------
        # Output file
        # -----------------------------------------

        output_path = (
            outputs_dir
            / create_unique_filename(
                ".mp4"
            )
        )

        # -----------------------------------------
        # FFmpeg binary
        # -----------------------------------------

        ffmpeg_binary = getattr(
            settings,
            "FFMPEG_BINARY",
            "ffmpeg",
        )

        # -----------------------------------------
        # Build FFmpeg command
        # -----------------------------------------

        command = build_crop_command(
            ffmpeg_binary,
            input_path,
            output_path,
            x,
            y,
            width,
            height,
        )

        logger.info(
            "Running FFmpeg crop."
        )

        logger.info(
            "Crop command: %s",
            " ".join(
                str(item)
                for item in command
            ),
        )

        # -----------------------------------------
        # Execute FFmpeg
        # -----------------------------------------

        run_ffmpeg(
            command
        )

        # -----------------------------------------
        # Verify output
        # -----------------------------------------

        if not output_path.exists():

            raise RuntimeError(
                "FFmpeg did not create output."
            )

        if output_path.stat().st_size <= 0:

            raise RuntimeError(
                "Generated video is empty."
            )

        logger.info(
            "Crop completed successfully."
        )

        logger.info(
            "Output: %s",
            output_path,
        )

        return output_path

    except ValueError:

        logger.warning(
            "Crop validation failed.",
            exc_info=True,
        )

        raise

    except Exception as exc:

        logger.exception(
            "Crop processing failed."
        )

        raise RuntimeError(
            "Unable to process the video."
        ) from exc