import os


ALLOWED_VIDEO_EXTENSIONS = {
    ".mp4",
    ".mov",
    ".avi",
    ".mkv",
    ".webm",
}

ALLOWED_IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
}


def validate_video(video):

    if not video:
        raise ValueError("Please upload a video.")

    extension = os.path.splitext(
        video.name
    )[1].lower()

    if extension not in ALLOWED_VIDEO_EXTENSIONS:
        raise ValueError(
            "Unsupported video format."
        )

    return video


def validate_watermark(image):

    if not image:
        raise ValueError(
            "Please upload a watermark image."
        )

    extension = os.path.splitext(
        image.name
    )[1].lower()

    if extension not in ALLOWED_IMAGE_EXTENSIONS:
        raise ValueError(
            "Watermark must be PNG, JPG or JPEG."
        )

    return image


def validate_opacity(opacity):

    opacity = int(opacity)

    if opacity < 10 or opacity > 100:
        raise ValueError(
            "Opacity must be between 10 and 100."
        )

    return opacity


def validate_scale(scale):

    scale = int(scale)

    if scale < 5 or scale > 100:
        raise ValueError(
            "Scale must be between 5 and 100."
        )

    return scale


def validate_margin(value):

    value = int(value)

    if value < 0:
        raise ValueError(
            "Margin cannot be negative."
        )

    return value


def validate_position(position):

    allowed = [
        "top_left",
        "top_right",
        "bottom_left",
        "bottom_right",
        "center",
    ]

    if position not in allowed:
        raise ValueError(
            "Invalid watermark position."
        )

    return position


def validate_inputs(data):

    return {

        "video": validate_video(
            data["video"]
        ),

        "watermark": validate_watermark(
            data["watermark"]
        ),

        "opacity": validate_opacity(
            data["opacity"]
        ),

        "scale": validate_scale(
            data["scale"]
        ),

        "margin_x": validate_margin(
            data["margin_x"]
        ),

        "margin_y": validate_margin(
            data["margin_y"]
        ),

        "position": validate_position(
            data["position"]
        ),

    }