import os
import subprocess
from django.conf import settings


class AudioMerger:

    def __init__(self, video_path, audio_path):

        self.video_path = video_path
        self.audio_path = audio_path

    # ---------------------------------------
    # Merge Original Audio + Uploaded Audio
    # ---------------------------------------
    def merge(self):

        output_dir = os.path.join(settings.MEDIA_ROOT, "output")

        os.makedirs(output_dir, exist_ok=True)

        output_video = os.path.join(
            output_dir,
            "merged_video.mp4"
        )

        command = [

            "ffmpeg",

            "-y",

            "-i", self.video_path,

            "-i", self.audio_path,

            "-filter_complex",

            "[0:a][1:a]amix=inputs=2:duration=first:dropout_transition=2[a]",

            "-map", "0:v",

            "-map", "[a]",

            "-c:v", "copy",

            "-c:a", "aac",

            output_video

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_video

    # ---------------------------------------
    # Merge With Volume Control
    # ---------------------------------------
    def merge_with_volume(
        self,
        original_volume=1.0,
        new_volume=1.0
    ):

        output_dir = os.path.join(settings.MEDIA_ROOT, "output")

        os.makedirs(output_dir, exist_ok=True)

        output_video = os.path.join(
            output_dir,
            "merged_volume_video.mp4"
        )

        filter_complex = (
            f"[0:a]volume={original_volume}[a0];"
            f"[1:a]volume={new_volume}[a1];"
            f"[a0][a1]amix=inputs=2:duration=first[a]"
        )

        command = [

            "ffmpeg",

            "-y",

            "-i", self.video_path,

            "-i", self.audio_path,

            "-filter_complex", filter_complex,

            "-map", "0:v",

            "-map", "[a]",

            "-c:v", "copy",

            "-c:a", "aac",

            output_video

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_video