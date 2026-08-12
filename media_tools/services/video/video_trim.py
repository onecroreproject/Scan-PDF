import logging
import os
from pathlib import Path
from uuid import uuid4

from django.conf import settings
from moviepy.editor import VideoFileClip


logger = logging.getLogger(__name__)


def trim_video(video_file, start_time, end_time):
    """
    Trim an uploaded video using MoviePy.

    Args:
        video_file: Django UploadedFile
        start_time: Start time in seconds.
        end_time: End time in seconds.

    Returns:
        str: Absolute path of the trimmed video.
    """

    input_path = None
    output_path = None
    clip = None
    trimmed_clip = None

    try:

        # --------------------------------------------------
        # Validate time values
        # --------------------------------------------------

        try:
            start_time = float(start_time)
            end_time = float(end_time)

        except (TypeError, ValueError):

            raise ValueError(
                "Start time and end time must be valid numbers."
            )

        if start_time < 0:

            raise ValueError(
                "Start time cannot be negative."
            )

        if end_time <= start_time:

            raise ValueError(
                "End time must be greater than start time."
            )

        logger.info(
            "Trim request: %s | start=%s | end=%s",
            video_file.name,
            start_time,
            end_time,
        )

        # --------------------------------------------------
        # Create temporary input file
        # --------------------------------------------------

        extension = (
            Path(video_file.name).suffix.lower()
            or ".mp4"
        )

        input_path = os.path.join(
            settings.MEDIA_ROOT,
            "media_tools",
            "temp",
            f"input_{uuid4().hex}{extension}",
        )

        os.makedirs(
            os.path.dirname(input_path),
            exist_ok=True,
        )

        with open(
            input_path,
            "wb",
        ) as destination:

            for chunk in video_file.chunks():
                destination.write(chunk)

        logger.info(
            "Temporary input created: %s",
            input_path,
        )

        # --------------------------------------------------
        # Open video
        # --------------------------------------------------

        clip = VideoFileClip(
            input_path
        )

        duration = float(
            clip.duration
        )

        logger.info(
            "Video duration: %.2f seconds",
            duration,
        )

        # --------------------------------------------------
        # Validate against actual duration
        # --------------------------------------------------

        if start_time >= duration:

            raise ValueError(
                f"Start time ({start_time:g}s) "
                f"must be less than the video duration "
                f"({duration:.2f}s)."
            )

        if end_time > duration:

            raise ValueError(
                f"End time ({end_time:g}s) "
                f"cannot exceed the video duration "
                f"({duration:.2f}s)."
            )

        # --------------------------------------------------
        # Create trimmed clip
        # --------------------------------------------------

        trimmed_clip = clip.subclip(
            start_time,
            end_time,
        )

        # --------------------------------------------------
        # Output directory
        # --------------------------------------------------

        output_directory = os.path.join(
            settings.MEDIA_ROOT,
            "media_tools",
            "trimmed_videos",
        )

        os.makedirs(
            output_directory,
            exist_ok=True,
        )

        output_path = os.path.join(
            output_directory,
            f"trimmed_{uuid4().hex}.mp4",
        )

        logger.info(
            "Creating trimmed video: %s",
            output_path,
        )

        # --------------------------------------------------
        # Write output
        # --------------------------------------------------

        trimmed_clip.write_videofile(
            output_path,
            codec="libx264",
            audio_codec="aac",
            temp_audiofile=os.path.join(
                output_directory,
                f"temp_audio_{uuid4().hex}.m4a",
            ),
            remove_temp=True,
            logger=None,
        )

        logger.info(
            "Trim completed successfully."
        )

        # --------------------------------------------------
        # Validate output
        # --------------------------------------------------

        if not os.path.exists(output_path):

            raise ValueError(
                "Trimmed video was not created."
            )

        output_size = os.path.getsize(
            output_path
        )

        if output_size <= 0:

            raise ValueError(
                "Generated video file is empty."
            )

        logger.info(
            "Output size: %s bytes",
            output_size,
        )

        return output_path

    except ValueError:

        logger.exception(
            "Video trimming validation failed."
        )

        if output_path and os.path.exists(
            output_path
        ):

            try:
                os.remove(output_path)
            except OSError:
                logger.warning(
                    "Could not remove failed output.",
                    exc_info=True,
                )

        raise

    except Exception as exc:

        logger.exception(
            "Video trimming failed."
        )

        if output_path and os.path.exists(
            output_path
        ):

            try:
                os.remove(output_path)
            except OSError:
                logger.warning(
                    "Could not remove failed output.",
                    exc_info=True,
                )

        raise ValueError(
            "The uploaded video could not be processed."
        ) from exc

    finally:

        # --------------------------------------------------
        # Close MoviePy clips FIRST
        # --------------------------------------------------

        if trimmed_clip is not None:

            try:
                trimmed_clip.close()
            except Exception:
                logger.warning(
                    "Could not close trimmed clip.",
                    exc_info=True,
                )

        if clip is not None:

            try:
                clip.close()
            except Exception:
                logger.warning(
                    "Could not close input clip.",
                    exc_info=True,
                )

        # --------------------------------------------------
        # Remove temporary input
        # --------------------------------------------------

        if input_path and os.path.exists(
            input_path
        ):

            try:
                os.remove(input_path)

            except OSError:

                logger.warning(
                    "Could not remove temporary input: %s",
                    input_path,
                    exc_info=True,
                )