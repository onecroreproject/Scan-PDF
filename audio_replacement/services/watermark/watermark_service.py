import os
import uuid
import shutil
import subprocess

from django.conf import settings

from .validation import validate_inputs
from .image_processor import WatermarkImageProcessor
from .ffmpeg_command import build_ffmpeg_command


def process_watermark_video(data):
    """
    Process video and add image watermark.
    """

    validated = validate_inputs(data)

    input_dir = os.path.join(
        settings.MEDIA_ROOT,
        "video_tools",
        "watermark",
        "input",
    )

    output_dir = os.path.join(
        settings.MEDIA_ROOT,
        "video_tools",
        "watermark",
        "output",
    )

    os.makedirs(
        input_dir,
        exist_ok=True,
    )

    os.makedirs(
        output_dir,
        exist_ok=True,
    )

    # --------------------------------
    # Save uploaded video
    # --------------------------------

    video_ext = os.path.splitext(
        validated["video"].name
    )[1]

    video_path = os.path.join(
        input_dir,
        f"{uuid.uuid4().hex}{video_ext}",
    )

    with open(
        video_path,
        "wb+",
    ) as destination:

        for chunk in validated["video"].chunks():
            destination.write(chunk)

    # --------------------------------
    # Save watermark image
    # --------------------------------

    image_ext = os.path.splitext(
        validated["watermark"].name
    )[1]

    image_path = os.path.join(
        input_dir,
        f"{uuid.uuid4().hex}{image_ext}",
    )

    with open(
        image_path,
        "wb+",
    ) as destination:

        for chunk in validated["watermark"].chunks():
            destination.write(chunk)

    # --------------------------------
    # Resize + Opacity
    # --------------------------------

    processor = WatermarkImageProcessor(
        image_path,
    )

    processed_image = processor.process(
        scale=validated["scale"],
        opacity=validated["opacity"],
    )

    # --------------------------------
    # Output file
    # --------------------------------

    output_video = os.path.join(
        output_dir,
        f"watermark_{uuid.uuid4().hex}.mp4",
    )

    # --------------------------------
    # FFmpeg
    # --------------------------------

    command = build_ffmpeg_command(
        input_video=video_path,
        watermark_image=processed_image,
        output_video=output_video,
        position=validated["position"],
        margin_x=validated["margin_x"],
        margin_y=validated["margin_y"],
    )

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        raise Exception(result.stderr)

    # --------------------------------
    # Cleanup
    # --------------------------------

    for file in [
        video_path,
        image_path,
        processed_image,
    ]:

        if os.path.exists(file):

            try:
                os.remove(file)

            except Exception:
                pass

    return output_video