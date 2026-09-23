import logging

from django.conf import settings

from media_tools.services.ffmpeg_commands import (
    build_flip_command,
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
    "gif",
    "mpeg",
    "flv",
    "wmv",
}


ALLOWED_FLIP_MODES = {
    "horizontal",
    "vertical",
}


def process_flip(
    video,
    flip_mode="horizontal",
    output_format="mp4",
):
    """
    Flip a video horizontally or vertically using FFmpeg.

    Parameters
    ----------
    video:
        Uploaded video file.

    flip_mode:
        Either:
            - horizontal
            - vertical

    output_format:
        Output video format.

    Returns
    -------
    pathlib.Path
        Path to the generated video.
    """

    input_path = None
    output_path = None

    try:

        # ==================================================
        # VALIDATION
        # ==================================================

        if not video:
            raise ValueError(
                "Video file is required."
            )


        if flip_mode not in ALLOWED_FLIP_MODES:
            raise ValueError(
                "Invalid flip mode."
            )


        if output_format not in ALLOWED_FORMATS:
            raise ValueError(
                "Invalid output format."
            )


        logger.info(
            "Starting video flip operation."
        )


        logger.info(
            "Flip settings: mode=%s, format=%s",
            flip_mode,
            output_format,
        )


        # ==================================================
        # SAVE UPLOADED VIDEO
        # ==================================================

        input_path = save_uploaded_file(
            video
        )


        logger.info(
            "Uploaded video saved: %s",
            input_path,
        )


        # ==================================================
        # READ VIDEO INFORMATION
        # ==================================================

        info = get_video_info(
            input_path
        )


        if not info:
            raise ValueError(
                "Unable to read video information."
            )


        source_width = info.get(
            "width",
            0,
        )

        source_height = info.get(
            "height",
            0,
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


        # ==================================================
        # OUTPUT DIRECTORY
        # ==================================================

        _, outputs_dir, _ = (
            get_video_directories()
        )


        # ==================================================
        # OUTPUT FILE
        # ==================================================

        extension = (
            f".{output_format}"
        )


        output_path = (
            outputs_dir
            / create_unique_filename(
                extension
            )
        )


        logger.info(
            "Output path: %s",
            output_path,
        )


        # ==================================================
        # FFMPEG BINARY
        # ==================================================

        ffmpeg_binary = getattr(
            settings,
            "FFMPEG_BINARY",
            "ffmpeg",
        )


        # ==================================================
        # BUILD COMMAND
        # ==================================================

        command = build_flip_command(
            ffmpeg_binary,
            input_path,
            output_path,
            flip_mode,
            output_format,
        )


        logger.info(
            "FFmpeg flip command created."
        )


        # ==================================================
        # RUN FFMPEG
        # ==================================================

        run_ffmpeg(
            command
        )


        # ==================================================
        # VERIFY OUTPUT
        # ==================================================

        if not output_path.exists():
            raise RuntimeError(
                "FFmpeg did not create output."
            )


        if output_path.stat().st_size <= 0:
            raise RuntimeError(
                "Generated video is empty."
            )


        logger.info(
            "Video flip completed successfully: %s",
            output_path,
        )


        return output_path


    # ======================================================
    # VALIDATION ERROR
    # ======================================================

    except ValueError:

        logger.warning(
            "Video flip validation failed."
        )

        raise


    # ======================================================
    # PROCESSING ERROR
    # ======================================================

    except Exception as exc:

        logger.exception(
            "Video flip processing failed."
        )

        raise RuntimeError(
            "Unable to flip the video."
        ) from exc
