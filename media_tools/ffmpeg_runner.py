import logging
import subprocess

from django.conf import settings


logger = logging.getLogger("media_tools")


class FFmpegError(Exception):
    """FFmpeg processing error."""


def run_ffmpeg(
    command,
    timeout=3600,
):
    """
    Execute FFmpeg safely.
    """

    try:
        if not command:
            raise FFmpegError(
                "FFmpeg command is empty."
            )

        ffmpeg_binary = getattr(
            settings,
            "FFMPEG_BINARY",
            "ffmpeg",
        )

        if command[0] != ffmpeg_binary:
            logger.warning(
                "Unexpected FFmpeg executable: %s",
                command[0],
            )

        logger.info(
            "Starting FFmpeg process."
        )

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=timeout,
            check=False,
        )

        if result.returncode != 0:
            logger.error(
                "FFmpeg failed. Code=%s",
                result.returncode,
            )

            logger.error(
                "FFmpeg stderr: %s",
                result.stderr[-4000:],
            )

            raise FFmpegError(
                "Video processing failed."
            )

        logger.info(
            "FFmpeg completed successfully."
        )

        return result

    except FFmpegError:
        raise

    except FileNotFoundError as exc:
        logger.exception(
            "FFmpeg executable not found."
        )
        raise FFmpegError(
            "FFmpeg is not installed or configured."
        ) from exc

    except subprocess.TimeoutExpired as exc:
        logger.exception(
            "FFmpeg processing timed out."
        )
        raise FFmpegError(
            "Video processing timed out."
        ) from exc

    except OSError as exc:
        logger.exception(
            "FFmpeg operating system error."
        )
        raise FFmpegError(
            "Unable to start video processing."
        ) from exc

    except Exception as exc:
        logger.exception(
            "Unexpected FFmpeg error."
        )
        raise FFmpegError(
            "Unable to process video."
        ) from exc