import logging

from django.conf import settings

from media_tools.services.ffmpeg_commands import (
    build_resize_command,
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


ALLOWED_FIT_MODES = {
    "fit",
    "fill",
}


def process_resize(
    video,
    width,
    height,
    aspect_ratio="",
    fit_mode="fit",
    zoom=1.0,
    position_x=0,
    position_y=0,
    background_color="#000000",
    output_format="mp4",
):
    """
    Resize and position a video using FFmpeg.
    """

    input_path = None
    output_path = None

    try:

        if not video:
            raise ValueError(
                "Video file is required."
            )

        if width < 2 or height < 2:
            raise ValueError(
                "Width and height must be at least 2."
            )

        if fit_mode not in ALLOWED_FIT_MODES:
            raise ValueError(
                "Invalid fit mode."
            )

        if output_format not in ALLOWED_FORMATS:
            raise ValueError(
                "Invalid output format."
            )

        if zoom < 0.1 or zoom > 5:
            raise ValueError(
                "Zoom must be between 0.1 and 5."
            )

        if not isinstance(position_x, int):
            raise ValueError(
                "Invalid horizontal position."
            )

        if not isinstance(position_y, int):
            raise ValueError(
                "Invalid vertical position."
            )

        if (
            not background_color
            or not background_color.startswith("#")
        ):
            raise ValueError(
                "Invalid background color."
            )

        logger.info(
            "Starting resize operation."
        )

        # Save uploaded video
        input_path = save_uploaded_file(
            video
        )

        # Read source video metadata
        info = get_video_info(
            input_path
        )

        if info["width"] <= 0:
            raise ValueError(
                "Invalid source video width."
            )

        if info["height"] <= 0:
            raise ValueError(
                "Invalid source video height."
            )

        logger.info(
            "Source dimensions: %sx%s",
            info["width"],
            info["height"],
        )

        logger.info(
            "Output dimensions: %sx%s",
            width,
            height,
        )

        logger.info(
            "Resize settings: format=%s, "
            "fit=%s, zoom=%s, x=%s, y=%s",
            output_format,
            fit_mode,
            zoom,
            position_x,
            position_y,
        )

        # Get directories
        _, outputs_dir, _ = (
            get_video_directories()
        )

        # Create output extension
        extension = (
            f".{output_format}"
        )

        output_path = (
            outputs_dir
            / create_unique_filename(
                extension
            )
        )

        ffmpeg_binary = getattr(
            settings,
            "FFMPEG_BINARY",
            "ffmpeg",
        )

        command = build_resize_command(
            ffmpeg_binary,
            input_path,
            output_path,
            width,
            height,
            fit_mode,
            zoom,
            position_x,
            position_y,
            background_color,
            output_format,
        )

        # Run FFmpeg
        run_ffmpeg(command)

        # Verify output
        if not output_path.exists():
            raise RuntimeError(
                "FFmpeg did not create output."
            )

        if output_path.stat().st_size <= 0:
            raise RuntimeError(
                "Generated video is empty."
            )

        logger.info(
            "Resize completed successfully: %s",
            output_path,
        )

        return output_path

    except ValueError:
        logger.warning(
            "Resize validation failed."
        )
        raise

    except Exception as exc:
        logger.exception(
            "Resize processing failed."
        )

        raise RuntimeError(
            "Unable to resize the video."
        ) from exc