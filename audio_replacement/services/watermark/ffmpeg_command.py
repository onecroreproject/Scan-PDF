import os


def get_position(position, margin_x, margin_y):

    positions = {

        "top_left": (
            str(margin_x),
            str(margin_y),
        ),

        "top_right": (
            f"main_w-overlay_w-{margin_x}",
            str(margin_y),
        ),

        "bottom_left": (
            str(margin_x),
            f"main_h-overlay_h-{margin_y}",
        ),

        "bottom_right": (
            f"main_w-overlay_w-{margin_x}",
            f"main_h-overlay_h-{margin_y}",
        ),

        "center": (
            "(main_w-overlay_w)/2",
            "(main_h-overlay_h)/2",
        ),

    }

    return positions.get(
        position,
        positions["bottom_right"],
    )


def build_ffmpeg_command(
    input_video,
    watermark_image,
    output_video,
    position,
    margin_x,
    margin_y,
):

    x, y = get_position(
        position,
        margin_x,
        margin_y,
    )

    command = [

        "ffmpeg",

        "-y",

        "-i",
        input_video,

        "-i",
        watermark_image,

        "-filter_complex",
        f"overlay={x}:{y}",

        "-codec:a",
        "copy",

        output_video,

    ]

    return command