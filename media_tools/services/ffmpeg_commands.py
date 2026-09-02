def build_crop_command(
    ffmpeg_binary,
    input_path,
    output_path,
    x,
    y,
    width,
    height,
):
    """
    Build FFmpeg command for video cropping.

    The crop values come from the realtime
    JavaScript crop editor.
    """

    try:
        x = int(x)
        y = int(y)
        width = int(width)
        height = int(height)

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

        crop_filter = (
            f"crop={width}:{height}:{x}:{y}"
        )

        command = [
            str(ffmpeg_binary),

            "-y",

            "-i",
            str(input_path),

            "-vf",
            crop_filter,

            # Video
            "-c:v",
            "libx264",

            "-preset",
            "medium",

            "-crf",
            "18",

            "-pix_fmt",
            "yuv420p",

            # Audio
            "-c:a",
            "aac",

            "-b:a",
            "192k",

            # MP4 browser compatibility
            "-movflags",
            "+faststart",

            str(output_path),
        ]

        return command

    except ValueError:
        raise

    except Exception as exc:
        logger.exception(
            "Failed to build crop command."
        )

        raise RuntimeError(
            "Unable to build FFmpeg crop command."
        ) from exc


def build_resize_command(
    ffmpeg_binary,
    input_path,
    output_path,
    width,
    height,
    fit_mode="fit",
    zoom=1.0,
    position_x=0,
    position_y=0,
    background_color="#000000",
    output_format="mp4",
):
    """
    Build FFmpeg command for video resizing
    and canvas positioning.
    """

    if fit_mode == "fill":
        scale_filter = (
            f"scale={width}:{height}:"
            "force_original_aspect_ratio=increase,"
            f"crop={width}:{height}:"
            "(iw-ow)/2+"
            f"({position_x}):"
            "(ih-oh)/2+"
            f"({position_y})"
        )

    else:
        scale_filter = (
            f"scale={width}:{height}:"
            "force_original_aspect_ratio=decrease,"
            f"pad={width}:{height}:"
            "(ow-iw)/2+"
            f"({position_x}):"
            "(oh-ih)/2+"
            f"({position_y}):"
            f"color={background_color}"
        )

    if zoom != 1.0:
        scale_filter = (
            f"scale=iw*{zoom}:ih*{zoom},"
            + scale_filter
        )

    if output_format == "gif":

        return [
            ffmpeg_binary,
            "-y",
            "-i",
            str(input_path),
            "-vf",
            scale_filter,
            "-an",
            str(output_path),
        ]

    return [
        ffmpeg_binary,
        "-y",
        "-i",
        str(input_path),
        "-vf",
        scale_filter,
        "-c:v",
        "libx264",
        "-preset",
        "medium",
        "-crf",
        "23",
        "-c:a",
        "aac",
        "-b:a",
        "128k",
        "-pix_fmt",
        "yuv420p",
        "-movflags",
        "+faststart",
        str(output_path),
    ]