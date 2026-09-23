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

import logging


logger = logging.getLogger("media_tools")


def build_trim_command(
    ffmpeg_binary,
    input_path,
    output_path,
    start_time,
    end_time,
    output_format="mp4",
):
    """
    Build FFmpeg command for extracting a selected
    section from a video.

    The video is re-encoded so the trim can start
    accurately at the selected timeline position.
    """

    try:

        start_time = float(
            start_time
        )

        end_time = float(
            end_time
        )

        if start_time < 0:
            raise ValueError(
                "Start time cannot be negative."
            )

        if end_time <= start_time:
            raise ValueError(
                "End time must be greater than start time."
            )

        duration = (
            end_time
            - start_time
        )

        if duration <= 0:
            raise ValueError(
                "Trim duration must be positive."
            )

        output_format = (
            str(output_format)
            .lower()
            .strip()
        )

        allowed_formats = {
            "mp4",
            "mov",
            "webm",
            "avi",
            "mkv",
            "m4v",
            "flv",
            "wmv",
            "gif",
            "mpeg",
        }

        if output_format not in allowed_formats:
            raise ValueError(
                "Unsupported output format."
            )

        # -----------------------------------------------------
        # GIF
        # -----------------------------------------------------
        #
        # GIF does not contain normal video/audio streams.
        # Generate a palette and use it for better quality.
        #
        # This requires a filter_complex command.
        # -----------------------------------------------------

        if output_format == "gif":

            filter_complex = (
                "[0:v]"
                "fps=15,"
                "scale="
                "min(720\\,iw):"
                "-1:flags=lanczos,"
                "split"
                "[a][b];"
                "[a]"
                "palettegen="
                "max_colors=256"
                "[p];"
                "[b][p]"
                "paletteuse="
                "dither=sierra2_4a"
            )

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                "-filter_complex",
                filter_complex,

                "-an",

                str(output_path),
            ]

        # -----------------------------------------------------
        # WEBM
        # -----------------------------------------------------

        if output_format == "webm":

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                # Video
                "-c:v",
                "libvpx-vp9",

                "-crf",
                "32",

                "-b:v",
                "0",

                "-pix_fmt",
                "yuv420p",

                # Audio
                "-c:a",
                "libopus",

                "-b:a",
                "128k",

                str(output_path),
            ]

        # -----------------------------------------------------
        # WMV
        # -----------------------------------------------------

        if output_format == "wmv":

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                "-c:v",
                "wmv2",

                "-b:v",
                "2M",

                "-c:a",
                "wmav2",

                "-b:a",
                "128k",

                str(output_path),
            ]

        # -----------------------------------------------------
        # FLV
        # -----------------------------------------------------

        if output_format == "flv":

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                "-c:v",
                "flv",

                "-b:v",
                "2M",

                "-c:a",
                "aac",

                "-b:a",
                "128k",

                str(output_path),
            ]

        # -----------------------------------------------------
        # MPEG
        # -----------------------------------------------------

        if output_format == "mpeg":

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                "-c:v",
                "mpeg2video",

                "-b:v",
                "4M",

                "-c:a",
                "mp2",

                "-b:a",
                "192k",

                str(output_path),
            ]

        # -----------------------------------------------------
        # AVI
        # -----------------------------------------------------

        if output_format == "avi":

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                "-c:v",
                "mpeg4",

                "-q:v",
                "4",

                "-c:a",
                "aac",

                "-b:a",
                "192k",

                str(output_path),
            ]

        # -----------------------------------------------------
        # MKV
        # -----------------------------------------------------

        if output_format == "mkv":

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                "-c:v",
                "libx264",

                "-preset",
                "medium",

                "-crf",
                "18",

                "-pix_fmt",
                "yuv420p",

                "-c:a",
                "aac",

                "-b:a",
                "192k",

                str(output_path),
            ]

        # -----------------------------------------------------
        # MOV
        # -----------------------------------------------------

        if output_format == "mov":

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                "-c:v",
                "libx264",

                "-preset",
                "medium",

                "-crf",
                "18",

                "-pix_fmt",
                "yuv420p",

                "-c:a",
                "aac",

                "-b:a",
                "192k",

                str(output_path),
            ]

        # -----------------------------------------------------
        # M4V
        # -----------------------------------------------------

        if output_format == "m4v":

            return [
                str(ffmpeg_binary),

                "-y",

                "-ss",
                f"{start_time:.3f}",

                "-i",
                str(input_path),

                "-t",
                f"{duration:.3f}",

                "-c:v",
                "libx264",

                "-preset",
                "medium",

                "-crf",
                "18",

                "-pix_fmt",
                "yuv420p",

                "-c:a",
                "aac",

                "-b:a",
                "192k",

                "-movflags",
                "+faststart",

                str(output_path),
            ]

        # -----------------------------------------------------
        # MP4
        # -----------------------------------------------------
        #
        # Default browser-friendly output.
        # -----------------------------------------------------

        return [
            str(ffmpeg_binary),

            "-y",

            "-ss",
            f"{start_time:.3f}",

            "-i",
            str(input_path),

            "-t",
            f"{duration:.3f}",

            "-c:v",
            "libx264",

            "-preset",
            "medium",

            "-crf",
            "18",

            "-pix_fmt",
            "yuv420p",

            "-c:a",
            "aac",

            "-b:a",
            "192k",

            "-movflags",
            "+faststart",

            str(output_path),
        ]

    except ValueError:
        raise

    except Exception as exc:

        logger.exception(
            "Failed to build trim FFmpeg command."
        )

        raise RuntimeError(
            "Unable to build FFmpeg trim command."
        ) from exc

def build_rotate_command(
    ffmpeg_binary,
    input_path,
    output_path,
    rotation,
    output_format="mp4",
):
    """
    Build an FFmpeg command for video rotation.

    rotation:
        90  = clockwise 90 degrees
        180 = 180 degrees
        270 = clockwise 270 degrees
              / counter-clockwise 90 degrees
    """

    rotation = str(rotation).strip()

    rotation_filters = {
        "90": "transpose=1",
        "180": "transpose=1,transpose=1",
        "270": "transpose=2",
    }

    if rotation not in rotation_filters:
        raise ValueError(
            "Invalid rotation angle."
        )

    video_filter = rotation_filters[
        rotation
    ]

    command = [
        ffmpeg_binary,
        "-y",
        "-i",
        str(input_path),
        "-vf",
        video_filter,
        "-map",
        "0:v:0",
        "-map",
        "0:a?",
        "-c:v",
        "libx264",
        "-preset",
        "medium",
        "-crf",
        "18",
        "-c:a",
        "aac",
        "-movflags",
        "+faststart",
        str(output_path),
    ]

    return command


def build_rotate_command(
    ffmpeg_binary,
    input_path,
    output_path,
    rotation,
    output_format="mp4",
):
    """
    Build an FFmpeg command for rotating a video.

    Supported rotation values:
        0
        90
        180
        270

    The function only builds the command.
    FFmpeg execution is handled by ffmpeg_runner.py.
    """

    allowed_rotations = {
        0,
        90,
        180,
        270,
    }

    allowed_formats = {
        "mp4",
        "mov",
        "webm",
        "avi",
        "mkv",
        "m4v",
        "gif",
    }

    if rotation not in allowed_rotations:
        raise ValueError(
            "Invalid rotation angle."
        )

    if output_format not in allowed_formats:
        raise ValueError(
            "Invalid output format."
        )

    if not ffmpeg_binary:
        raise ValueError(
            "FFmpeg binary is required."
        )

    if not input_path:
        raise ValueError(
            "Input video path is required."
        )

    if not output_path:
        raise ValueError(
            "Output video path is required."
        )

    # --------------------------------------------------
    # Rotation filters
    # --------------------------------------------------

    rotation_filters = {
        0: "null",
        90: "transpose=1",
        180: "transpose=1,transpose=1",
        270: "transpose=2",
    }

    video_filter = rotation_filters[rotation]

    # --------------------------------------------------
    # Base command
    # --------------------------------------------------

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i",
        str(input_path),
        "-vf",
        video_filter,
    ]

    # --------------------------------------------------
    # Output settings
    # --------------------------------------------------

    if output_format == "mp4":
        command.extend(
            [
                "-c:v",
                "libx264",
                "-preset",
                "medium",
                "-crf",
                "23",
                "-c:a",
                "aac",
                "-movflags",
                "+faststart",
            ]
        )

    elif output_format == "mov":
        command.extend(
            [
                "-c:v",
                "libx264",
                "-preset",
                "medium",
                "-crf",
                "23",
                "-c:a",
                "aac",
            ]
        )

    elif output_format == "webm":
        command.extend(
            [
                "-c:v",
                "libvpx-vp9",
                "-crf",
                "30",
                "-b:v",
                "0",
                "-c:a",
                "libopus",
            ]
        )

    elif output_format == "avi":
        command.extend(
            [
                "-c:v",
                "mpeg4",
                "-q:v",
                "5",
                "-c:a",
                "mp3",
            ]
        )

    elif output_format == "mkv":
        command.extend(
            [
                "-c:v",
                "libx264",
                "-preset",
                "medium",
                "-crf",
                "23",
                "-c:a",
                "aac",
            ]
        )

    elif output_format == "m4v":
        command.extend(
            [
                "-c:v",
                "libx264",
                "-preset",
                "medium",
                "-crf",
                "23",
                "-c:a",
                "aac",
            ]
        )

    elif output_format == "gif":
        command.extend(
            [
                "-an",
                "-f",
                "gif",
            ]
        )

    # --------------------------------------------------
    # Output path
    # --------------------------------------------------

    command.append(
        str(output_path)
    )

    return command


def build_flip_command(
    ffmpeg_binary,
    input_path,
    output_path,
    flip_mode="horizontal",
    output_format="mp4",
):
    """
    Build FFmpeg command for flipping video horizontally or vertically.
    """
    video_filter = "hflip" if flip_mode == "horizontal" else "vflip"
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-vf", video_filter,
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-c:a", "copy",
        "-movflags", "+faststart",
        str(output_path),
    ]
    return command


def build_speed_command(
    ffmpeg_binary,
    input_path,
    output_path,
    speed=1.0,
    keep_audio=True,
    output_format="mp4",
):
    """
    Build FFmpeg command for video playback speed adjustment.
    """
    try:
        speed = float(speed)
        if speed <= 0:
            speed = 1.0
    except (ValueError, TypeError):
        speed = 1.0

    pts_factor = 1.0 / speed
    video_filter = f"setpts={pts_factor:.4f}*PTS"

    # Handle audio tempo
    # atempo filter supports values between 0.5 and 2.0
    atempo_filters = []
    temp_speed = speed
    while temp_speed > 2.0:
        atempo_filters.append("atempo=2.0")
        temp_speed /= 2.0
    while temp_speed < 0.5:
        atempo_filters.append("atempo=0.5")
        temp_speed /= 0.5
    atempo_filters.append(f"atempo={temp_speed:.4f}")
    audio_filter = ",".join(atempo_filters)

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-filter:v", video_filter,
    ]

    if keep_audio:
        command.extend(["-filter:a", audio_filter])
    else:
        command.append("-an")

    command.extend([
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-movflags", "+faststart",
        str(output_path),
    ])
    return command


def build_compress_command(
    ffmpeg_binary,
    input_path,
    output_path,
    compression_level="medium",
    output_format="mp4",
):
    """
    Build FFmpeg command for intelligent video compression.
    """
    # Level to CRF mapping
    crf_map = {
        "low": "24",     # High Quality
        "medium": "28",  # Balanced
        "high": "34",    # Maximum Compression
    }
    crf = crf_map.get(compression_level, "28")
    preset = "slow" if compression_level == "high" else "medium"

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-c:v", "libx264",
        "-preset", preset,
        "-crf", crf,
        "-pix_fmt", "yuv420p",
        "-c:a", "aac",
        "-b:a", "128k",
        "-movflags", "+faststart",
        str(output_path),
    ]
    return command


def build_convert_command(
    ffmpeg_binary,
    input_path,
    output_path,
    target_format="mp4",
):
    """
    Build FFmpeg command for cross-format video conversion.
    """
    target_format = str(target_format).lower().replace(".", "")
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
    ]

    if target_format in ("mp4", "m4v", "mov"):
        command.extend([
            "-c:v", "libx264",
            "-preset", "medium",
            "-crf", "22",
            "-pix_fmt", "yuv420p",
            "-c:a", "aac",
            "-b:a", "192k",
            "-movflags", "+faststart",
        ])
    elif target_format == "webm":
        command.extend([
            "-c:v", "libvpx-vp9",
            "-crf", "30",
            "-b:v", "0",
            "-c:a", "libopus",
            "-b:a", "128k",
        ])
    elif target_format == "avi":
        command.extend([
            "-c:v", "mpeg4",
            "-q:v", "4",
            "-c:a", "mp3",
            "-b:a", "192k",
        ])
    elif target_format == "mkv":
        command.extend([
            "-c:v", "libx264",
            "-preset", "medium",
            "-crf", "22",
            "-c:a", "aac",
        ])
    elif target_format == "gif":
        command.extend([
            "-vf", "fps=15,scale=480:-1:flags=lanczos",
            "-an",
        ])
    else:
        command.extend([
            "-c:v", "libx264",
            "-c:a", "aac",
        ])

    command.append(str(output_path))
    return command


def build_video_to_gif_command(
    ffmpeg_binary,
    input_path,
    output_path,
    fps=15,
    width=480,
):
    """
    Build two-pass high-quality GIF generation command with palettegen/use.
    """
    width = int(width) if width else 480
    scale_str = f"min({width}\\,iw):-1:flags=lanczos"
    filter_complex = (
        f"[0:v]fps={fps},scale={scale_str},split[a][b];"
        f"[a]palettegen=max_colors=256[p];"
        f"[b][p]paletteuse=dither=sierra2_4a"
    )
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-filter_complex", filter_complex,
        "-an",
        str(output_path),
    ]
    return command


def build_gif_to_video_command(
    ffmpeg_binary,
    input_path,
    output_path,
    output_format="mp4",
):
    """
    Convert GIF to universal playable MP4 video.
    """
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-movflags", "faststart",
        "-pix_fmt", "yuv420p",
        "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "20",
        str(output_path),
    ]
    return command


def build_remove_audio_command(
    ffmpeg_binary,
    input_path,
    output_path,
    output_format="mp4",
):
    """
    Build FFmpeg command to mute/strip audio track.
    """
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-c:v", "copy",
        "-an",
        str(output_path),
    ]
    return command


def build_extract_audio_command(
    ffmpeg_binary,
    input_path,
    output_path,
    audio_format="mp3",
    bitrate="192k",
):
    """
    Extract audio track from video to pure audio file.
    """
    audio_format = str(audio_format).lower().replace(".", "")
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-vn",
    ]
    if audio_format == "mp3":
        command.extend(["-c:a", "libmp3lame", "-b:a", bitrate])
    elif audio_format in ("m4a", "aac"):
        command.extend(["-c:a", "aac", "-b:a", bitrate])
    elif audio_format == "wav":
        command.extend(["-c:a", "pcm_s16le"])
    elif audio_format == "ogg":
        command.extend(["-c:a", "libvorbis", "-b:a", bitrate])
    else:
        command.extend(["-c:a", "copy"])

    command.append(str(output_path))
    return command


def build_volume_command(
    ffmpeg_binary,
    input_path,
    output_path,
    volume_ratio=1.0,
    output_format="mp4",
):
    """
    Adjust audio volume of video.
    """
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-c:v", "copy",
        "-filter:a", f"volume={volume_ratio:.2f}",
        "-c:a", "aac",
        "-b:a", "192k",
        str(output_path),
    ]
    return command


def build_reverse_command(
    ffmpeg_binary,
    input_path,
    output_path,
    reverse_audio=True,
    output_format="mp4",
):
    """
    Reverse video and optionally audio playback.
    """
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-vf", "reverse",
    ]
    if reverse_audio:
        command.extend(["-af", "areverse"])
    else:
        command.append("-an")

    command.extend([
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-c:a", "aac",
        "-movflags", "+faststart",
        str(output_path),
    ])
    return command


def build_fade_command(
    ffmpeg_binary,
    input_path,
    output_path,
    fade_in_duration=1.0,
    fade_out_duration=1.0,
    total_duration=10.0,
    color="black",
    output_format="mp4",
):
    """
    Apply fade-in and/or fade-out transitions to video and audio.
    """
    fade_in_duration = max(0.0, float(fade_in_duration))
    fade_out_duration = max(0.0, float(fade_out_duration))
    total_duration = max(1.0, float(total_duration))

    v_filters = []
    a_filters = []

    if fade_in_duration > 0:
        v_filters.append(f"fade=t=in:st=0:d={fade_in_duration:.2f}:c={color}")
        a_filters.append(f"afade=t=in:ss=0:d={fade_in_duration:.2f}")

    if fade_out_duration > 0:
        out_start = max(0.0, total_duration - fade_out_duration)
        v_filters.append(f"fade=t=out:st={out_start:.2f}:d={fade_out_duration:.2f}:c={color}")
        a_filters.append(f"afade=t=out:st={out_start:.2f}:d={fade_out_duration:.2f}")

    vf_str = ",".join(v_filters) if v_filters else "null"
    af_str = ",".join(a_filters) if a_filters else "anull"

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-vf", vf_str,
        "-af", af_str,
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-c:a", "aac",
        "-movflags", "+faststart",
        str(output_path),
    ]
    return command


def build_color_filter_command(
    ffmpeg_binary,
    input_path,
    output_path,
    brightness=0.0,
    contrast=1.0,
    saturation=1.0,
    filter_type="none",
    output_format="mp4",
):
    """
    Apply brightness, contrast, saturation, or artistic filters (grayscale, sepia, invert).
    """
    v_filters = []

    if filter_type == "grayscale":
        v_filters.append("hue=s=0")
    elif filter_type == "sepia":
        v_filters.append("colorchannelmixer=.393:.769:.189:0:.349:.686:.168:0:.272:.534:.131")
    elif filter_type == "invert":
        v_filters.append("negate")

    # Add EQ adjustments
    v_filters.append(f"eq=brightness={brightness:.2f}:contrast={contrast:.2f}:saturation={saturation:.2f}")

    vf_str = ",".join(v_filters)
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-vf", vf_str,
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-c:a", "copy",
        "-movflags", "+faststart",
        str(output_path),
    ]
    return command


def build_fps_command(
    ffmpeg_binary,
    input_path,
    output_path,
    fps=30,
    output_format="mp4",
):
    """
    Change video frame rate (FPS).
    """
    fps = int(fps)
    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-filter:v", f"fps={fps}",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-c:a", "copy",
        "-movflags", "+faststart",
        str(output_path),
    ]
    return command


def build_loop_command(
    ffmpeg_binary,
    input_path,
    output_path,
    loop_count=2,
    output_format="mp4",
):
    """
    Loop a video seamlessly N times.
    """
    loop_count = max(2, int(loop_count))
    command = [
        str(ffmpeg_binary),
        "-y",
        "-stream_loop", str(loop_count - 1),
        "-i", str(input_path),
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-c:a", "aac",
        "-movflags", "+faststart",
        str(output_path),
    ]
    return command


def build_frame_extract_command(
    ffmpeg_binary,
    input_path,
    output_path,
    time_offset=0.0,
):
    """
    Extract a single pristine still frame from video at timestamp.
    """
    time_offset = max(0.0, float(time_offset))
    command = [
        str(ffmpeg_binary),
        "-y",
        "-ss", f"{time_offset:.3f}",
        "-i", str(input_path),
        "-vframes", "1",
        "-q:v", "2",
        str(output_path),
    ]
    return command


def build_watermark_command(
    ffmpeg_binary,
    input_path,
    watermark_path,
    output_path,
    position="bottom-right",
    opacity=0.8,
    scale_percent=15,
    output_format="mp4",
):
    """
    Overlay image watermark onto video with customizable position and opacity.
    """
    scale_factor = max(0.05, min(0.5, float(scale_percent) / 100.0))
    opacity = max(0.1, min(1.0, float(opacity)))

    # Position coordinates
    pos_map = {
        "top-left": "20:20",
        "top-right": "W-w-20:20",
        "bottom-left": "20:H-h-20",
        "bottom-right": "W-w-20:H-h-20",
        "center": "(W-w)/2:(H-h)/2",
    }
    pos_coords = pos_map.get(position, "W-w-20:H-h-20")

    filter_complex = (
        f"[1:v]scale=iw*{scale_factor:.3f}:-1,format=rgba,"
        f"colorchannelmixer=aa={opacity:.2f}[wm];"
        f"[0:v][wm]overlay={pos_coords}"
    )

    command = [
        str(ffmpeg_binary),
        "-y",
        "-i", str(input_path),
        "-i", str(watermark_path),
        "-filter_complex", filter_complex,
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "22",
        "-pix_fmt", "yuv420p",
        "-c:a", "copy",
        "-movflags", "+faststart",
        str(output_path),
    ]
    return command

