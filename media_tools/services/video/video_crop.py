# media_tools/services/video/video_crop.py

import logging
import os
import uuid
from fractions import Fraction

import av
import numpy as np

from django.conf import settings


logger = logging.getLogger(__name__)


# ============================================================
# SUPPORTED FORMATS
# ============================================================

SUPPORTED_VIDEO_EXTENSIONS = {
    ".mp4",
    ".mov",
    ".avi",
    ".mkv",
    ".webm",
}


# ============================================================
# PRESET ASPECT RATIOS
# ============================================================

ASPECT_RATIOS = {
    "1:1": (1, 1),
    "2:2": (1, 1),

    "4:3": (4, 3),
    "3:4": (3, 4),

    "16:9": (16, 9),
    "9:16": (9, 16),

    "3:2": (3, 2),
    "2:3": (2, 3),

    "5:4": (5, 4),
    "4:5": (4, 5),

    "8:5": (8, 5),
    "5:8": (5, 8),

    "3:7": (3, 7),
    "7:3": (7, 3),
}


# ============================================================
# GENERAL HELPERS
# ============================================================

def _validate_input_video(input_path):
    """
    Validate input video before processing.
    """

    if not input_path:
        raise ValueError(
            "No video file was provided."
        )

    if not os.path.isfile(input_path):
        raise ValueError(
            "The uploaded video could not be found."
        )

    extension = os.path.splitext(
        input_path
    )[1].lower()

    if extension not in SUPPORTED_VIDEO_EXTENSIONS:
        raise ValueError(
            "Unsupported video format. "
            "Supported formats are MP4, MOV, AVI, MKV and WEBM."
        )

    file_size = os.path.getsize(
        input_path
    )

    if file_size <= 0:
        raise ValueError(
            "The uploaded video is empty."
        )

    logger.info(
        "Input video validated: %s",
        input_path,
    )


def _get_output_directory():
    """
    Create permanent crop output directory.
    """

    output_directory = os.path.join(
        settings.MEDIA_ROOT,
        "media_tools",
        "cropped_videos",
    )

    os.makedirs(
        output_directory,
        exist_ok=True,
    )

    return output_directory


def _make_even(value):
    """
    Video encoders commonly require even dimensions.
    """

    value = int(value)

    if value < 2:
        value = 2

    if value % 2 != 0:
        value -= 1

    return value


# ============================================================
# VIDEO INFORMATION
# ============================================================

def _get_video_information(input_path):
    """
    Read video information using PyAV.

    Returns:
        width
        height
        fps
        time_base
        codec_name
    """

    container = None

    try:

        container = av.open(
            input_path,
            mode="r",
        )

        video_stream = next(
            (
                stream
                for stream in container.streams
                if stream.type == "video"
            ),
            None,
        )

        if video_stream is None:
            raise ValueError(
                "No video stream was found."
            )

        width = int(
            video_stream.codec_context.width
            or video_stream.width
        )

        height = int(
            video_stream.codec_context.height
            or video_stream.height
        )

        if width <= 0 or height <= 0:
            raise ValueError(
                "Invalid video dimensions."
            )

        # ---------------------------------------------
        # FPS
        # ---------------------------------------------

        fps = video_stream.average_rate

        if fps is None:
            fps = video_stream.base_rate

        if fps is None:
            fps = Fraction(30, 1)

        else:
            fps = Fraction(
                fps.numerator,
                fps.denominator,
            )

        # ---------------------------------------------
        # Time base
        # ---------------------------------------------

        time_base = video_stream.time_base

        if time_base is None:
            time_base = Fraction(1, 90000)

        else:
            time_base = Fraction(
                time_base.numerator,
                time_base.denominator,
            )

        codec_name = (
            video_stream.codec_context.name
            or "unknown"
        )

        logger.info(
            "Video information: "
            "width=%s height=%s fps=%s "
            "time_base=%s codec=%s",
            width,
            height,
            fps,
            time_base,
            codec_name,
        )

        return {
            "width": width,
            "height": height,
            "fps": fps,
            "time_base": time_base,
            "codec": codec_name,
        }

    except av.error.FFmpegError as exc:

        logger.exception(
            "PyAV could not read the video."
        )

        raise ValueError(
            "The uploaded video could not be read."
        ) from exc

    except Exception:

        logger.exception(
            "Unexpected error while reading video information."
        )

        raise

    finally:

        if container is not None:

            try:
                container.close()

            except Exception:
                logger.warning(
                    "Could not close input container.",
                    exc_info=True,
                )


# ============================================================
# PRESET CROP
# ============================================================

def _calculate_ratio_crop(
    source_width,
    source_height,
    ratio,
):
    """
    Calculate centered crop rectangle
    for a selected aspect ratio.
    """

    if ratio not in ASPECT_RATIOS:
        raise ValueError(
            f"Unsupported crop ratio: {ratio}"
        )

    ratio_width, ratio_height = (
        ASPECT_RATIOS[ratio]
    )

    target_ratio = Fraction(
        ratio_width,
        ratio_height,
    )

    source_ratio = Fraction(
        source_width,
        source_height,
    )

    # --------------------------------------------------------
    # Source wider than target
    # --------------------------------------------------------

    if source_ratio > target_ratio:

        crop_height = source_height

        crop_width = int(
            crop_height * target_ratio
        )

        crop_x = (
            source_width - crop_width
        ) // 2

        crop_y = 0

    # --------------------------------------------------------
    # Source taller than target
    # --------------------------------------------------------

    else:

        crop_width = source_width

        crop_height = int(
            crop_width / target_ratio
        )

        crop_x = 0

        crop_y = (
            source_height - crop_height
        ) // 2

    crop_width = _make_even(
        crop_width
    )

    crop_height = _make_even(
        crop_height
    )

    crop_x = max(
        0,
        int(crop_x),
    )

    crop_y = max(
        0,
        int(crop_y),
    )

    # Final boundary protection.

    if crop_x + crop_width > source_width:

        crop_width = (
            source_width - crop_x
        )

    if crop_y + crop_height > source_height:

        crop_height = (
            source_height - crop_y
        )

    crop_width = _make_even(
        crop_width
    )

    crop_height = _make_even(
        crop_height
    )

    return (
        crop_x,
        crop_y,
        crop_width,
        crop_height,
    )


# ============================================================
# CUSTOM CROP VALIDATION
# ============================================================

def _validate_crop_coordinates(
    source_width,
    source_height,
    crop_x,
    crop_y,
    crop_width,
    crop_height,
):
    """
    Validate crop rectangle received from frontend.
    """

    try:

        crop_x = int(
            round(float(crop_x))
        )

        crop_y = int(
            round(float(crop_y))
        )

        crop_width = int(
            round(float(crop_width))
        )

        crop_height = int(
            round(float(crop_height))
        )

    except (
        TypeError,
        ValueError,
    ) as exc:

        raise ValueError(
            "Invalid crop coordinates."
        ) from exc

    if crop_x < 0:
        raise ValueError(
            "Crop X cannot be negative."
        )

    if crop_y < 0:
        raise ValueError(
            "Crop Y cannot be negative."
        )

    if crop_width <= 0:
        raise ValueError(
            "Crop width must be greater than zero."
        )

    if crop_height <= 0:
        raise ValueError(
            "Crop height must be greater than zero."
        )

    if crop_x >= source_width:
        raise ValueError(
            "Crop X is outside the video."
        )

    if crop_y >= source_height:
        raise ValueError(
            "Crop Y is outside the video."
        )

    # Keep rectangle inside video.

    crop_width = min(
        crop_width,
        source_width - crop_x,
    )

    crop_height = min(
        crop_height,
        source_height - crop_y,
    )

    crop_width = _make_even(
        crop_width
    )

    crop_height = _make_even(
        crop_height
    )

    if crop_width < 2:
        raise ValueError(
            "Crop width is too small."
        )

    if crop_height < 2:
        raise ValueError(
            "Crop height is too small."
        )

    return (
        crop_x,
        crop_y,
        crop_width,
        crop_height,
    )


# ============================================================
# CUSTOM WIDTH / HEIGHT
# ============================================================

def _calculate_custom_crop(
    source_width,
    source_height,
    custom_width,
    custom_height,
):
    """
    Create centered custom crop.
    """

    try:

        custom_width = int(
            custom_width
        )

        custom_height = int(
            custom_height
        )

    except (
        TypeError,
        ValueError,
    ) as exc:

        raise ValueError(
            "Custom width and height must be numbers."
        ) from exc

    if custom_width <= 0:

        raise ValueError(
            "Custom width must be greater than zero."
        )

    if custom_height <= 0:

        raise ValueError(
            "Custom height must be greater than zero."
        )

    if custom_width > source_width:

        raise ValueError(
            "Custom width cannot be larger "
            "than the original video width."
        )

    if custom_height > source_height:

        raise ValueError(
            "Custom height cannot be larger "
            "than the original video height."
        )

    custom_width = _make_even(
        custom_width
    )

    custom_height = _make_even(
        custom_height
    )

    crop_x = (
        source_width - custom_width
    ) // 2

    crop_y = (
        source_height - custom_height
    ) // 2

    return (
        crop_x,
        crop_y,
        custom_width,
        custom_height,
    )


# ============================================================
# FRAME CROP
# ============================================================

def _crop_frame(
    frame,
    crop_x,
    crop_y,
    crop_width,
    crop_height,
):
    """
    Crop a PyAV VideoFrame using NumPy.

    IMPORTANT:
    VideoFrame.crop() does not exist in PyAV.

    Therefore:

        VideoFrame
            ↓
        NumPy ndarray
            ↓
        NumPy slicing
            ↓
        VideoFrame.from_ndarray()

    """

    try:

        # Convert PyAV frame to ndarray.
        frame_array = frame.to_ndarray(
            format="rgb24"
        )

        frame_height, frame_width = (
            frame_array.shape[:2]
        )

        # ----------------------------------------------------
        # Safety validation
        # ----------------------------------------------------

        if crop_x < 0:
            crop_x = 0

        if crop_y < 0:
            crop_y = 0

        crop_x = min(
            crop_x,
            frame_width - 1,
        )

        crop_y = min(
            crop_y,
            frame_height - 1,
        )

        crop_width = min(
            crop_width,
            frame_width - crop_x,
        )

        crop_height = min(
            crop_height,
            frame_height - crop_y,
        )

        crop_width = _make_even(
            crop_width
        )

        crop_height = _make_even(
            crop_height
        )

        if crop_width <= 0:
            raise ValueError(
                "Calculated crop width is invalid."
            )

        if crop_height <= 0:
            raise ValueError(
                "Calculated crop height is invalid."
            )

        # ----------------------------------------------------
        # NumPy crop
        # ----------------------------------------------------

        cropped_array = frame_array[
            crop_y:crop_y + crop_height,
            crop_x:crop_x + crop_width,
        ]

        if cropped_array.size == 0:

            raise ValueError(
                "The selected crop area is empty."
            )

        # ----------------------------------------------------
        # Convert back to PyAV
        # ----------------------------------------------------

        cropped_frame = (
            av.VideoFrame.from_ndarray(
                cropped_array,
                format="rgb24",
            )
        )

        return cropped_frame

    except Exception as exc:

        logger.exception(
            "Failed to crop video frame."
        )

        raise ValueError(
            "The video frame could not be processed."
        ) from exc


# ============================================================
# VIDEO ENCODER
# ============================================================

def _create_output_stream(
    output_container,
    input_stream,
    crop_width,
    crop_height,
):
    """
    Create output H.264 video stream.

    All Fraction values are kept as Fraction objects
    to avoid PyAV's `.numerator` errors.
    """

    # --------------------------------------------------------
    # FPS
    # --------------------------------------------------------

    fps = input_stream.average_rate

    if fps is None:
        fps = input_stream.base_rate

    if fps is None:
        fps = Fraction(30, 1)

    else:
        fps = Fraction(
            fps.numerator,
            fps.denominator,
        )

    # --------------------------------------------------------
    # Output stream
    # --------------------------------------------------------

    output_stream = output_container.add_stream(
        "libx264",
        rate=fps,
    )

    # IMPORTANT:
    # Never assign a float to framerate.

    output_stream.width = _make_even(
        crop_width
    )

    output_stream.height = _make_even(
        crop_height
    )

    output_stream.pix_fmt = "yuv420p"

    # Quality / speed.
    output_stream.options = {
        "preset": "veryfast",
        "crf": "20",
    }

    # --------------------------------------------------------
    # Time base
    # --------------------------------------------------------

    output_stream.time_base = Fraction(
        1,
        90000,
    )

    logger.info(
        "Output stream created: "
        "width=%s height=%s fps=%s",
        output_stream.width,
        output_stream.height,
        fps,
    )

    return output_stream


# ============================================================
# MAIN PYAV PROCESSING
# ============================================================

def _process_video(
    input_path,
    output_path,
    crop_x,
    crop_y,
    crop_width,
    crop_height,
):
    """
    Process video using PyAV + NumPy.
    """

    input_container = None
    output_container = None

    try:

        # ----------------------------------------------------
        # Open input
        # ----------------------------------------------------

        input_container = av.open(
            input_path,
            mode="r",
        )

        input_stream = next(
            (
                stream
                for stream in input_container.streams
                if stream.type == "video"
            ),
            None,
        )

        if input_stream is None:

            raise ValueError(
                "No video stream was found."
            )

        # ----------------------------------------------------
        # Open output
        # ----------------------------------------------------

        output_container = av.open(
            output_path,
            mode="w",
        )

        output_stream = _create_output_stream(
            output_container,
            input_stream,
            crop_width,
            crop_height,
        )

        # ----------------------------------------------------
        # Process frames
        # ----------------------------------------------------

        frame_count = 0

        for frame in input_container.decode(
            input_stream
        ):

            frame_count += 1

            cropped_frame = _crop_frame(
                frame,
                crop_x,
                crop_y,
                crop_width,
                crop_height,
            )

            # ------------------------------------------------
            # Preserve timing safely.
            # ------------------------------------------------

            if frame.pts is not None:

                try:

                    cropped_frame.pts = (
                        frame.pts
                    )

                except Exception:

                    logger.debug(
                        "Could not preserve frame PTS.",
                        exc_info=True,
                    )

            # Do NOT assign None as time_base.

            try:

                if frame.time_base is not None:

                    cropped_frame.time_base = (
                        frame.time_base
                    )

            except Exception:

                logger.debug(
                    "Could not preserve frame time_base.",
                    exc_info=True,
                )

            # ------------------------------------------------
            # Encode
            # ------------------------------------------------

            for packet in output_stream.encode(
                cropped_frame
            ):

                output_container.mux(
                    packet
                )

            # ------------------------------------------------
            # Progress logging
            # ------------------------------------------------

            if frame_count % 100 == 0:

                logger.info(
                    "Crop processing: %s frames processed.",
                    frame_count,
                )

        # ----------------------------------------------------
        # Flush encoder
        # ----------------------------------------------------

        for packet in output_stream.encode():

            output_container.mux(
                packet
            )

        logger.info(
            "Video crop processing completed. "
            "Frames processed: %s",
            frame_count,
        )

    except av.error.FFmpegError as exc:

        logger.exception(
            "PyAV video processing failed."
        )

        raise ValueError(
            "The video could not be processed."
        ) from exc

    except ValueError:

        logger.warning(
            "Video crop validation/processing failed.",
            exc_info=True,
        )

        raise

    except Exception as exc:

        logger.exception(
            "Unexpected video crop processing error."
        )

        raise ValueError(
            "The video could not be processed."
        ) from exc

    finally:

        if input_container is not None:

            try:
                input_container.close()

            except Exception:

                logger.warning(
                    "Could not close input video container.",
                    exc_info=True,
                )

        if output_container is not None:

            try:
                output_container.close()

            except Exception:

                logger.warning(
                    "Could not close output video container.",
                    exc_info=True,
                )


# ============================================================
# PUBLIC SERVICE FUNCTION
# ============================================================

def crop_video(
    input_path,
    ratio=None,
    custom_crop=False,
    custom_width=None,
    custom_height=None,
    crop_x=None,
    crop_y=None,
    crop_width=None,
    crop_height=None,
):
    """
    Public video crop service.

    Supported modes:

    1. Preset ratio
    2. Custom width/height
    3. Frontend draggable crop rectangle

    Returns:

        {
            "output_path": "...",
            "output_url": "...",
            "crop_x": ...,
            "crop_y": ...,
            "crop_width": ...,
            "crop_height": ...,
            "source_width": ...,
            "source_height": ...,
        }
    """

    output_path = None

    try:

        logger.info(
            "================================================"
        )

        logger.info(
            "Starting video crop."
        )

        logger.info(
            "Input: %s",
            input_path,
        )

        logger.info(
            "Ratio: %s | custom_crop=%s",
            ratio,
            custom_crop,
        )

        # ----------------------------------------------------
        # Validate input
        # ----------------------------------------------------

        _validate_input_video(
            input_path
        )

        # ----------------------------------------------------
        # Get video information
        # ----------------------------------------------------

        video_info = _get_video_information(
            input_path
        )

        source_width = video_info[
            "width"
        ]

        source_height = video_info[
            "height"
        ]

        # ----------------------------------------------------
        # Decide crop mode
        # ----------------------------------------------------

        # FRONTEND DRAG CROP
        if (
            crop_x is not None
            and crop_y is not None
            and crop_width is not None
            and crop_height is not None
        ):

            logger.info(
                "Crop mode: frontend rectangle."
            )

            (
                crop_x,
                crop_y,
                crop_width,
                crop_height,
            ) = _validate_crop_coordinates(
                source_width,
                source_height,
                crop_x,
                crop_y,
                crop_width,
                crop_height,
            )

        # CUSTOM WIDTH / HEIGHT
        elif custom_crop:

            logger.info(
                "Crop mode: custom dimensions."
            )

            (
                crop_x,
                crop_y,
                crop_width,
                crop_height,
            ) = _calculate_custom_crop(
                source_width,
                source_height,
                custom_width,
                custom_height,
            )

        # PRESET RATIO
        elif ratio and ratio != "free":

            logger.info(
                "Crop mode: preset ratio=%s",
                ratio,
            )

            (
                crop_x,
                crop_y,
                crop_width,
                crop_height,
            ) = _calculate_ratio_crop(
                source_width,
                source_height,
                ratio,
            )

        else:

            raise ValueError(
                "Please select a crop ratio "
                "or crop area."
            )

        # ----------------------------------------------------
        # Final validation
        # ----------------------------------------------------

        (
            crop_x,
            crop_y,
            crop_width,
            crop_height,
        ) = _validate_crop_coordinates(
            source_width,
            source_height,
            crop_x,
            crop_y,
            crop_width,
            crop_height,
        )

        logger.info(
            "Final crop rectangle: "
            "x=%s y=%s width=%s height=%s",
            crop_x,
            crop_y,
            crop_width,
            crop_height,
        )

        # ----------------------------------------------------
        # Output directory
        # ----------------------------------------------------

        output_directory = (
            _get_output_directory()
        )

        output_filename = (
            f"cropped_"
            f"{uuid.uuid4().hex}.mp4"
        )

        output_path = os.path.join(
            output_directory,
            output_filename,
        )

        # ----------------------------------------------------
        # Process
        # ----------------------------------------------------

        _process_video(
            input_path=input_path,
            output_path=output_path,
            crop_x=crop_x,
            crop_y=crop_y,
            crop_width=crop_width,
            crop_height=crop_height,
        )

        # ----------------------------------------------------
        # Verify output
        # ----------------------------------------------------

        if not os.path.exists(
            output_path
        ):

            raise ValueError(
                "Crop completed but output file "
                "was not created."
            )

        output_size = os.path.getsize(
            output_path
        )

        if output_size <= 0:

            raise ValueError(
                "The cropped video is empty."
            )

        # ----------------------------------------------------
        # Build media URL
        # ----------------------------------------------------

        relative_path = os.path.relpath(
            output_path,
            settings.MEDIA_ROOT,
        )

        relative_path = relative_path.replace(
            os.sep,
            "/",
        )

        output_url = (
            settings.MEDIA_URL.rstrip("/")
            + "/"
            + relative_path
        )

        logger.info(
            "Video crop successful."
        )

        logger.info(
            "Output: %s",
            output_path,
        )

        logger.info(
            "Output URL: %s",
            output_url,
        )

        logger.info(
            "Output size: %s bytes",
            output_size,
        )

        logger.info(
            "================================================"
        )

        return {
            "output_path": output_path,
            "output_url": output_url,

            "crop_x": crop_x,
            "crop_y": crop_y,

            "crop_width": crop_width,
            "crop_height": crop_height,

            "source_width": source_width,
            "source_height": source_height,

            "ratio": ratio,
        }

    except ValueError:

        logger.warning(
            "Video crop validation/processing failed.",
            exc_info=True,
        )

        # Delete incomplete output.

        if (
            output_path
            and os.path.exists(output_path)
        ):

            try:

                os.remove(
                    output_path
                )

            except OSError:

                logger.warning(
                    "Could not remove failed crop output: %s",
                    output_path,
                    exc_info=True,
                )

        raise

    except Exception as exc:

        logger.exception(
            "Unexpected error during video crop."
        )

        if (
            output_path
            and os.path.exists(output_path)
        ):

            try:

                os.remove(
                    output_path
                )

            except OSError:

                logger.warning(
                    "Could not remove failed output.",
                    exc_info=True,
                )

        raise ValueError(
            "The video could not be cropped."
        ) from exc