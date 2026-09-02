import json
import logging
import subprocess

from django.conf import settings


logger = logging.getLogger("media_tools")


class MediaInfoError(Exception):
    """Media information error."""


def get_video_info(video_path):
    """
    Get video information using FFprobe.
    """

    try:
        ffprobe = getattr(
            settings,
            "FFPROBE_BINARY",
            "ffprobe",
        )

        command = [
            ffprobe,
            "-v",
            "error",
            "-print_format",
            "json",
            "-show_streams",
            "-show_format",
            str(video_path),
        ]

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120,
            check=False,
        )

        if result.returncode != 0:
            logger.error(
                "FFprobe failed: %s",
                result.stderr[-2000:],
            )

            raise MediaInfoError(
                "Unable to read video information."
            )

        try:
            data = json.loads(
                result.stdout
            )
        except json.JSONDecodeError as exc:
            logger.exception(
                "Invalid FFprobe JSON."
            )
            raise MediaInfoError(
                "Unable to read video information."
            ) from exc

        video_stream = next(
            (
                stream
                for stream
                in data.get("streams", [])
                if stream.get("codec_type")
                == "video"
            ),
            None,
        )

        if not video_stream:
            raise MediaInfoError(
                "No video stream found."
            )

        width = int(
            video_stream.get(
                "width",
                0,
            )
        )

        height = int(
            video_stream.get(
                "height",
                0,
            )
        )

        duration = float(
            video_stream.get(
                "duration",
                data.get(
                    "format",
                    {},
                ).get(
                    "duration",
                    0,
                ),
            )
            or 0
        )

        fps = parse_fps(
            video_stream.get(
                "r_frame_rate"
            )
        )

        return {
            "width": width,
            "height": height,
            "duration": duration,
            "fps": fps,
            "video_codec": video_stream.get(
                "codec_name"
            ),
            "audio_codec": get_audio_codec(
                data.get("streams", [])
            ),
        }

    except MediaInfoError:
        raise

    except FileNotFoundError as exc:
        logger.exception(
            "FFprobe executable not found."
        )
        raise MediaInfoError(
            "FFprobe is not installed or configured."
        ) from exc

    except subprocess.TimeoutExpired as exc:
        logger.exception(
            "FFprobe timed out."
        )
        raise MediaInfoError(
            "Reading video information timed out."
        ) from exc

    except OSError as exc:
        logger.exception(
            "FFprobe operating system error."
        )
        raise MediaInfoError(
            "Unable to read video information."
        ) from exc

    except Exception as exc:
        logger.exception(
            "Unexpected media information error."
        )
        raise MediaInfoError(
            "Unable to read video information."
        ) from exc


def parse_fps(value):
    try:
        if not value or value == "0/0":
            return 0

        numerator, denominator = (
            value.split("/")
        )

        denominator = float(denominator)

        if denominator == 0:
            return 0

        return (
            float(numerator)
            / denominator
        )

    except (
        ValueError,
        TypeError,
        ZeroDivisionError,
    ):
        return 0


def get_audio_codec(streams):
    try:
        for stream in streams:
            if (
                stream.get("codec_type")
                == "audio"
            ):
                return stream.get(
                    "codec_name"
                )

        return None

    except Exception:
        logger.exception(
            "Unable to read audio codec."
        )
        return None