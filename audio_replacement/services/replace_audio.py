import os
import subprocess
from django.conf import settings


class AudioReplacer:

    def __init__(self, video_path, audio_path):

        self.video_path = video_path
        self.audio_path = audio_path

    # ---------------------------------------
    # Replace Original Audio
    # ---------------------------------------
    def replace(self):

        output_dir = os.path.join(settings.MEDIA_ROOT, "output")

        os.makedirs(output_dir, exist_ok=True)

        output_video = os.path.join(
            output_dir,
            "output_video.mp4"
        )

        command = [

            "ffmpeg",

            "-y",

            "-i",
            self.video_path,

            "-i",
            self.audio_path,

            "-map",
            "0:v:0",

            "-map",
            "1:a:0",

            "-c:v",
            "copy",

            "-c:a",
            "aac",

            "-shortest",

            output_video

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_video

    # ---------------------------------------
    # Replace with Custom Output Name
    # ---------------------------------------
    def replace_as(self, filename):

        output_dir = os.path.join(settings.MEDIA_ROOT, "output")

        os.makedirs(output_dir, exist_ok=True)

        output_video = os.path.join(
            output_dir,
            filename
        )

        command = [

            "ffmpeg",

            "-y",

            "-i",
            self.video_path,

            "-i",
            self.audio_path,

            "-map",
            "0:v:0",

            "-map",
            "1:a:0",

            "-c:v",
            "copy",

            "-c:a",
            "aac",

            "-shortest",

            output_video

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_video