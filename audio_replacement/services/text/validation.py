"""
validation.py

Purpose:
    Validate all Add Text to Video settings before
    passing them to the FFmpeg command builder.
"""


ALLOWED_FONTS = [
    "Arial",
    "Verdana",
    "Tahoma",
    "Georgia",
    "Times New Roman",
    "Courier New",
]


ALLOWED_COLORS = [
    "white",
    "black",
    "red",
    "green",
    "blue",
    "yellow",
]


ALLOWED_POSITIONS = [
    "center",
    "top",
    "bottom",
    "left",
    "right",
    "top_left",
    "top_right",
    "bottom_left",
    "bottom_right",
]





def validate_color(color):
    """
    Validate text color.
    """

    if color not in ALLOWED_COLORS:
        raise ValueError(
            f"Unsupported color: {color}"
        )

    return color


def validate_position(position):
    """
    Validate text position.
    """

    if position not in ALLOWED_POSITIONS:
        raise ValueError(
            f"Unsupported position: {position}"
        )

    return position


def validate_font_size(size):
    """
    Font size must be between 10 and 100.
    """

    size = int(size)

    if size < 10 or size > 100:
        raise ValueError(
            "Font size must be between 10 and 100."
        )

    return size


def validate_margin(value):
    """
    Margin cannot be negative.
    """

    value = int(value)

    if value < 0:
        raise ValueError(
            "Margin cannot be negative."
        )

    return value


def validate_opacity(opacity):
    """
    Opacity must be between 0 and 100.
    """

    opacity = int(opacity)

    if opacity < 0 or opacity > 100:
        raise ValueError(
            "Opacity must be between 0 and 100."
        )

    return opacity


def validate_duration(duration):
    """
    Duration must be greater than zero.
    """

    duration = int(duration)

    if duration <= 0:
        raise ValueError(
            "Duration must be greater than zero."
        )

    return duration


def validate_text(text):
    """
    Validate watermark text.
    """

    text = text.strip()

    if not text:
        raise ValueError(
            "Text cannot be empty."
        )

    if len(text) > 200:
        raise ValueError(
            "Maximum text length is 200 characters."
        )

    return text


def validate_inputs(data):
    """
    Validate all user inputs.

    Returns:
        Dictionary containing validated values.
    """

    return {

        "video": data["video"],

        "text": validate_text(
            data["text"]
        ),

       

        "font_size": validate_font_size(
            data["font_size"]
        ),

        "font_color": validate_color(
            data["font_color"]
        ),

        "position": validate_position(
            data["position"]
        ),

        "margin_x": validate_margin(
            data["margin_x"]
        ),

        "margin_y": validate_margin(
            data["margin_y"]
        ),

        "opacity": validate_opacity(
            data["opacity"]
        ),

        "duration": validate_duration(
            data["duration"]
        ),

    }