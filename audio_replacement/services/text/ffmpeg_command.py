from pathlib import Path

DEFAULT_FONT = Path("C:/Windows/Fonts/arial.ttf").as_posix()
DEFAULT_FONT = DEFAULT_FONT.replace(":", "\\:")


def get_position(position, margin_x, margin_y):

    positions = {

        "center": (
            "(w-text_w)/2",
            "(h-text_h)/2"
        ),

        "top": (
            "(w-text_w)/2",
            str(margin_y)
        ),

        "bottom": (
            "(w-text_w)/2",
            f"h-text_h-{margin_y}"
        ),

        "left": (
            str(margin_x),
            "(h-text_h)/2"
        ),

        "right": (
            f"w-text_w-{margin_x}",
            "(h-text_h)/2"
        ),

        "top_left": (
            str(margin_x),
            str(margin_y)
        ),

        "top_right": (
            f"w-text_w-{margin_x}",
            str(margin_y)
        ),

        "bottom_left": (
            str(margin_x),
            f"h-text_h-{margin_y}"
        ),

        "bottom_right": (
            f"w-text_w-{margin_x}",
            f"h-text_h-{margin_y}"
        ),

    }

    return positions.get(position, positions["center"])


def build_ffmpeg_command(
    input_video,
    output_video,
    text,
    font_size,
    font_color,
    position,
    margin_x,
    margin_y,
    opacity,
    duration,
):

    x, y = get_position(
        position,
        margin_x,
        margin_y,
    )

    alpha = opacity / 100

    text = (
        text.replace("\\", "\\\\")
            .replace(":", r"\:")
            .replace("'", r"\'")
    )

    draw_text = (
        f"drawtext="
        f"fontfile='{DEFAULT_FONT}':"
        f"text='{text}':"
        f"fontsize={font_size}:"
        f"fontcolor={font_color}@{alpha}:"
        f"x={x}:"
        f"y={y}:"
        f"enable='between(t,0,{duration})'"
    )
    print("=" * 80)
    print(draw_text)
    print("=" * 80)
    command = [
        "ffmpeg",
        "-y",
        "-i",
        input_video,
        "-vf",
        draw_text,
        "-c:a",
        "copy",
        output_video,
    ]

    return command