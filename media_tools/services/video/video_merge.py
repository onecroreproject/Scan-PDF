import logging
import os
from fractions import Fraction
from pathlib import Path
from uuid import uuid4

import av
from django.conf import settings


logger = logging.getLogger(__name__)


SUPPORTED_VIDEO_EXTENSIONS = {
    ".mp4",
    ".mov",
    ".mkv",
    ".avi",
    ".webm",
    ".m4v",
}


# ============================================================
# SAVE UPLOADED VIDEO
# ============================================================

def _save_uploaded_video(video_file):
    """
    Save one uploaded video to a unique temporary location.
    """

    if not video_file:
        raise ValueError("Invalid uploaded video.")

    extension = Path(
        video_file.name
    ).suffix.lower()

    if extension not in SUPPORTED_VIDEO_EXTENSIONS:
        raise ValueError(
            f"Unsupported video format: {extension}"
        )

    temp_directory = os.path.join(
        settings.MEDIA_ROOT,
        "media_tools",
        "merge_temp",
    )

    os.makedirs(
        temp_directory,
        exist_ok=True,
    )

    filename = (
        f"input_{uuid4().hex}{extension}"
    )

    file_path = os.path.join(
        temp_directory,
        filename,
    )

    try:

        with open(
            file_path,
            "wb",
        ) as destination:

            for chunk in video_file.chunks():
                destination.write(chunk)

        if not os.path.exists(file_path):
            raise ValueError(
                "Uploaded video could not be saved."
            )

        if os.path.getsize(file_path) <= 0:
            raise ValueError(
                "Uploaded video is empty."
            )

        logger.info(
            "Temporary video saved: %s",
            file_path,
        )

        return file_path

    except Exception:

        logger.exception(
            "Failed to save uploaded video."
        )

        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except OSError:
            logger.warning(
                "Could not remove failed upload: %s",
                file_path,
                exc_info=True,
            )

        raise


# ============================================================
# VIDEO INFORMATION
# ============================================================

def _get_video_info(file_path):
    """
    Read basic video information.
    """

    container = None

    try:

        container = av.open(
            file_path
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
                f"No video stream found in "
                f"{Path(file_path).name}"
            )

        width = int(
            video_stream.width
        )

        height = int(
            video_stream.height
        )

        frame_rate = (
            video_stream.average_rate
        )

        if frame_rate:

            fps = Fraction(
                frame_rate.numerator,
                frame_rate.denominator,
            )

        else:

            fps = Fraction(
                30,
                1,
            )

        # Avoid unreasonable FPS values.
        if fps <= 0:
            fps = Fraction(30, 1)

        if fps > 120:
            fps = Fraction(30, 1)

        return {
            "width": width,
            "height": height,
            "fps": fps,
            "has_audio": any(
                stream.type == "audio"
                for stream in container.streams
            ),
        }

    except Exception:

        logger.exception(
            "Could not inspect video: %s",
            file_path,
        )

        raise

    finally:

        if container is not None:

            try:
                container.close()
            except Exception:
                logger.warning(
                    "Could not close video container.",
                    exc_info=True,
                )


# ============================================================
# OUTPUT SIZE
# ============================================================

def _get_output_size(video_paths):
    """
    Use the first video resolution.

    H.264 yuv420p requires even dimensions.
    """

    first_info = _get_video_info(
        video_paths[0]
    )

    width = first_info["width"]
    height = first_info["height"]

    if width % 2:
        width -= 1

    if height % 2:
        height -= 1

    if width < 2 or height < 2:
        raise ValueError(
            "Invalid video dimensions."
        )

    return width, height


# ============================================================
# VIDEO MERGING
# ============================================================

def _merge_with_reencoding(
    video_paths,
    output_path,
):
    """
    Merge multiple videos using PyAV.

    Output:

        Video:
            H.264
            yuv420p
            fixed FPS

        Audio:
            AAC
            stereo
            48000 Hz

    Video and audio are processed for each source
    video instead of processing all video first and
    all audio afterwards.
    """

    output_container = None

    try:

        # ----------------------------------------------------
        # Output settings
        # ----------------------------------------------------

        output_width, output_height = (
            _get_output_size(
                video_paths
            )
        )

        first_info = _get_video_info(
            video_paths[0]
        )

        output_fps = first_info["fps"]

        # Explicit time bases.
        video_time_base = Fraction(
            1,
            output_fps,
        )

        audio_rate = 48000

        audio_time_base = Fraction(
            1,
            audio_rate,
        )

        logger.info(
            "Merge output: %sx%s @ %s FPS",
            output_width,
            output_height,
            output_fps,
        )

        # ----------------------------------------------------
        # Open output container
        # ----------------------------------------------------

        output_container = av.open(
            output_path,
            mode="w",
        )

        # ----------------------------------------------------
        # VIDEO OUTPUT STREAM
        # ----------------------------------------------------

        output_video_stream = (
            output_container.add_stream(
                "libx264",
                rate=output_fps,
            )
        )

        output_video_stream.width = (
            output_width
        )

        output_video_stream.height = (
            output_height
        )

        output_video_stream.pix_fmt = (
            "yuv420p"
        )

        output_video_stream.time_base = (
            video_time_base
        )

        output_video_stream.options = {
            "preset": "ultrafast",
            "crf": "23",
        }

        # ----------------------------------------------------
        # AUDIO OUTPUT STREAM
        # ----------------------------------------------------

        output_audio_stream = (
            output_container.add_stream(
                "aac",
                rate=audio_rate,
            )
        )

        output_audio_stream.layout = (
            "stereo"
        )

        output_audio_stream.time_base = (
            audio_time_base
        )

        # ----------------------------------------------------
        # GLOBAL TIMESTAMPS
        # ----------------------------------------------------

        global_video_pts = 0
        global_audio_samples = 0

        # ----------------------------------------------------
        # PROCESS EACH VIDEO
        # ----------------------------------------------------

        for video_index, video_path in enumerate(
            video_paths,
            start=1,
        ):

            logger.info(
                "Processing video %d/%d: %s",
                video_index,
                len(video_paths),
                Path(video_path).name,
            )

            input_container = None

            # Start position of this video.
            segment_video_start = (
                global_video_pts
            )

            segment_audio_start = (
                global_audio_samples
            )

            try:

                input_container = av.open(
                    video_path
                )

                input_video_stream = next(
                    (
                        stream
                        for stream in input_container.streams
                        if stream.type == "video"
                    ),
                    None,
                )

                input_audio_stream = next(
                    (
                        stream
                        for stream in input_container.streams
                        if stream.type == "audio"
                    ),
                    None,
                )

                if input_video_stream is None:
                    raise ValueError(
                        f"No video stream found in "
                        f"{Path(video_path).name}"
                    )

                # ------------------------------------------------
                # AUDIO RESAMPLER
                # ------------------------------------------------

                resampler = None

                if input_audio_stream is not None:

                    resampler = (
                        av.audio.resampler.AudioResampler(
                            format="fltp",
                            layout="stereo",
                            rate=audio_rate,
                        )
                    )

                # ------------------------------------------------
                # DECODE INPUT IN PACKET ORDER
                #
                # This keeps video/audio processing together.
                # ------------------------------------------------

                streams_to_decode = [
                    input_video_stream
                ]

                if input_audio_stream is not None:
                    streams_to_decode.append(
                        input_audio_stream
                    )

                for packet in input_container.demux(
                    streams_to_decode
                ):

                    # ============================================
                    # VIDEO PACKET
                    # ============================================

                    if (
                        packet.stream.type
                        == "video"
                    ):

                        for frame in packet.decode():

                            # Normalize frame.
                            frame = frame.reformat(
                                width=output_width,
                                height=output_height,
                                format="yuv420p",
                            )

                            # Explicit video timestamp.
                            frame.pts = (
                                global_video_pts
                            )

                            # IMPORTANT:
                            # Do NOT use output stream time_base
                            # because PyAV may return None before
                            # stream initialization.

                            frame.time_base = (
                                video_time_base
                            )

                            packets = (
                                output_video_stream.encode(
                                    frame
                                )
                            )

                            for encoded_packet in packets:

                                output_container.mux(
                                    encoded_packet
                                )

                            global_video_pts += 1

                    # ============================================
                    # AUDIO PACKET
                    # ============================================

                    elif (
                        packet.stream.type
                        == "audio"
                    ):

                        if resampler is None:
                            continue

                        for frame in packet.decode():

                            converted_frames = (
                                resampler.resample(
                                    frame
                                )
                            )

                            if converted_frames is None:
                                continue

                            if not isinstance(
                                converted_frames,
                                list,
                            ):
                                converted_frames = [
                                    converted_frames
                                ]

                            for audio_frame in (
                                converted_frames
                            ):

                                if audio_frame is None:
                                    continue

                                audio_frame.pts = (
                                    global_audio_samples
                                )

                                audio_frame.time_base = (
                                    audio_time_base
                                )

                                packets = (
                                    output_audio_stream.encode(
                                        audio_frame
                                    )
                                )

                                for encoded_packet in packets:

                                    output_container.mux(
                                        encoded_packet
                                    )

                                if audio_frame.samples:
                                    global_audio_samples += (
                                        audio_frame.samples
                                    )

                # ------------------------------------------------
                # FLUSH AUDIO RESAMPLER
                # ------------------------------------------------

                if resampler is not None:

                    try:

                        remaining_frames = (
                            resampler.resample(
                                None
                            )
                        )

                        if remaining_frames is None:
                            remaining_frames = []

                        if not isinstance(
                            remaining_frames,
                            list,
                        ):
                            remaining_frames = [
                                remaining_frames
                            ]

                        for audio_frame in (
                            remaining_frames
                        ):

                            if audio_frame is None:
                                continue

                            audio_frame.pts = (
                                global_audio_samples
                            )

                            audio_frame.time_base = (
                                audio_time_base
                            )

                            packets = (
                                output_audio_stream.encode(
                                    audio_frame
                                )
                            )

                            for encoded_packet in packets:

                                output_container.mux(
                                    encoded_packet
                                )

                            if audio_frame.samples:
                                global_audio_samples += (
                                    audio_frame.samples
                                )

                    except Exception:

                        logger.warning(
                            "Audio resampler flush failed "
                            "for %s.",
                            Path(video_path).name,
                            exc_info=True,
                        )

                logger.info(
                    "Video %d completed. "
                    "Video frames=%d, audio samples=%d",
                    video_index,
                    global_video_pts
                    - segment_video_start,
                    global_audio_samples
                    - segment_audio_start,
                )

            finally:

                if input_container is not None:

                    try:
                        input_container.close()
                    except Exception:
                        logger.warning(
                            "Could not close input "
                            "container.",
                            exc_info=True,
                        )

        # ----------------------------------------------------
        # FLUSH VIDEO ENCODER
        # ----------------------------------------------------

        logger.info(
            "Flushing video encoder."
        )

        for encoded_packet in (
            output_video_stream.encode()
        ):

            output_container.mux(
                encoded_packet
            )

        # ----------------------------------------------------
        # FLUSH AUDIO ENCODER
        # ----------------------------------------------------

        logger.info(
            "Flushing audio encoder."
        )

        for encoded_packet in (
            output_audio_stream.encode()
        ):

            output_container.mux(
                encoded_packet
            )

        logger.info(
            "Video merge encoding completed."
        )

    except Exception:

        logger.exception(
            "PyAV video merge failed."
        )

        raise

    finally:

        if output_container is not None:

            try:
                output_container.close()
            except Exception:
                logger.warning(
                    "Could not close output "
                    "container.",
                    exc_info=True,
                )


# ============================================================
# PUBLIC SERVICE
# ============================================================

def merge_videos(video_files):
    """
    Public service used by Django views.

    Input:
        Multiple uploaded video files.

    Output:
        Absolute path of merged video.
    """

    input_paths = []
    output_path = None

    try:

        # ----------------------------------------------------
        # VALIDATION
        # ----------------------------------------------------

        if not video_files:

            raise ValueError(
                "Please upload videos."
            )

        video_files = list(
            video_files
        )

        if len(video_files) < 2:

            raise ValueError(
                "Please upload at least two videos."
            )

        logger.info(
            "Starting merge. "
            "Number of videos: %d",
            len(video_files),
        )

        # ----------------------------------------------------
        # SAVE FILES
        # ----------------------------------------------------

        for video_file in video_files:

            saved_path = (
                _save_uploaded_video(
                    video_file
                )
            )

            input_paths.append(
                saved_path
            )

        # ----------------------------------------------------
        # VALIDATE INPUT VIDEOS
        # ----------------------------------------------------

        for path in input_paths:

            info = _get_video_info(
                path
            )

            logger.info(
                "Input video: %s | "
                "Resolution=%sx%s | "
                "FPS=%s | Audio=%s",
                Path(path).name,
                info["width"],
                info["height"],
                info["fps"],
                info["has_audio"],
            )

        # ----------------------------------------------------
        # OUTPUT DIRECTORY
        # ----------------------------------------------------

        output_directory = os.path.join(
            settings.MEDIA_ROOT,
            "media_tools",
            "merged_videos",
        )

        os.makedirs(
            output_directory,
            exist_ok=True,
        )

        # ----------------------------------------------------
        # UNIQUE OUTPUT
        # ----------------------------------------------------

        output_filename = (
            f"merged_{uuid4().hex}.mp4"
        )

        output_path = os.path.join(
            output_directory,
            output_filename,
        )

        # ----------------------------------------------------
        # MERGE
        # ----------------------------------------------------

        _merge_with_reencoding(
            input_paths,
            output_path,
        )

        # ----------------------------------------------------
        # VERIFY OUTPUT
        # ----------------------------------------------------

        if not os.path.exists(
            output_path
        ):

            raise ValueError(
                "Merged video was not created."
            )

        output_size = os.path.getsize(
            output_path
        )

        if output_size <= 0:

            raise ValueError(
                "Merged video is empty."
            )

        # Try opening the generated MP4.
        test_container = None

        try:

            test_container = av.open(
                output_path
            )

            video_stream = next(
                (
                    stream
                    for stream in test_container.streams
                    if stream.type == "video"
                ),
                None,
            )

            if video_stream is None:

                raise ValueError(
                    "Generated file does not contain "
                    "a valid video stream."
                )

        finally:

            if test_container is not None:

                try:
                    test_container.close()
                except Exception:
                    pass

        logger.info(
            "Video merge successful: %s "
            "(%d bytes)",
            output_path,
            output_size,
        )

        return output_path

    except ValueError:

        logger.exception(
            "Video merge validation/processing "
            "error."
        )

        raise

    except Exception as exc:

        logger.exception(
            "Unexpected video merge error."
        )

        raise ValueError(
            "The uploaded videos could not be merged."
        ) from exc

    finally:

        # ----------------------------------------------------
        # CLEAN TEMPORARY UPLOADS
        # ----------------------------------------------------

        for path in input_paths:

            if not path:
                continue

            if not os.path.exists(path):
                continue

            try:

                os.remove(path)

                logger.info(
                    "Temporary file removed: %s",
                    path,
                )

            except OSError:

                logger.warning(
                    "Could not remove temporary file: %s",
                    path,
                    exc_info=True,
                )