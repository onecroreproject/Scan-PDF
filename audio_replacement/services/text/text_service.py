"""
text_service.py

Purpose:
    Process Add Text To Video request.
"""

import os
import uuid
import subprocess

from django.conf import settings

from .validation import validate_inputs
from .ffmpeg_command import build_ffmpeg_command


UPLOAD_FOLDER = os.path.join(
    settings.MEDIA_ROOT,
    "video_tools",
    "text",
    "input"
)


OUTPUT_FOLDER = os.path.join(
    settings.MEDIA_ROOT,
    "video_tools",
    "text",
    "output"
)


os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)

os.makedirs(
    OUTPUT_FOLDER,
    exist_ok=True
)


def process_text_video(data):
    """
    Main service function.
    """

    validated = validate_inputs(data)

    uploaded_video = validated["video"]

    extension = os.path.splitext(
        uploaded_video.name
    )[1]

    input_filename = f"{uuid.uuid4().hex}{extension}"

    input_path = os.path.join(
        UPLOAD_FOLDER,
        input_filename
    )

    with open(input_path, "wb+") as destination:

        for chunk in uploaded_video.chunks():

            destination.write(chunk)

    output_filename = (
        f"text_{uuid.uuid4().hex}.mp4"
    )

    output_path = os.path.join(
        OUTPUT_FOLDER,
        output_filename
    )

    command = build_ffmpeg_command(

        input_video=input_path,

        output_video=output_path,

        text=validated["text"],

        

        font_size=validated["font_size"],

        font_color=validated["font_color"],

        position=validated["position"],

        margin_x=validated["margin_x"],

        margin_y=validated["margin_y"],

        opacity=validated["opacity"],

        duration=validated["duration"]

    )

    result = subprocess.run(
        command,
        capture_output=True,
        text=True
    )

    print("========== FFMPEG STDOUT ==========")
    print(result.stdout)

    print("========== FFMPEG STDERR ==========")
    print(result.stderr)

    print("========== RETURN CODE ==========")
    print(result.returncode)

    if result.returncode != 0:
        raise Exception(result.stderr)
    return output_path