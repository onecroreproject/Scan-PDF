import os
import subprocess
from django.conf import settings


class DurationController:

    def __init__(self, media_path):
        self.media_path = media_path

    # ---------------------------------------
    # Trim Audio
    # start_time : HH:MM:SS or seconds
    # end_time   : HH:MM:SS or seconds
    # ---------------------------------------
    def trim_audio(self, start_time, end_time):

        output_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(output_dir, exist_ok=True)

        output_audio = os.path.join(
            output_dir,
            "trimmed_audio.wav"
        )

        command = [

            "ffmpeg",

            "-y",

            "-i",
            self.media_path,

            "-ss",
            str(start_time),

            "-to",
            str(end_time),

            "-c:a",
            "pcm_s16le",

            output_audio

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_audio

    # ---------------------------------------
    # Trim Video
    # ---------------------------------------
    def trim_video(self, start_time, end_time):

        output_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(output_dir, exist_ok=True)

        output_video = os.path.join(
            output_dir,
            "trimmed_video.mp4"
        )

        command = [

            "ffmpeg",

            "-y",

            "-i",
            self.media_path,

            "-ss",
            str(start_time),

            "-to",
            str(end_time),

            "-c:v",
            "copy",

            "-c:a",
            "copy",

            output_video

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_video

    # ---------------------------------------
    # Cut First N Seconds
    # ---------------------------------------
    def cut_first_seconds(self, seconds):

        return self.trim_audio(seconds, 999999)

    # ---------------------------------------
    # Keep First N Seconds
    # ---------------------------------------
    def keep_first_seconds(self, seconds):

        return self.trim_audio(0, seconds)

    # ---------------------------------------
    # Keep Last N Seconds (Audio)
    # ---------------------------------------
    def keep_last_seconds(self, seconds):

        command = [

            "ffprobe",

            "-v", "error",

            "-show_entries",
            "format=duration",

            "-of",
            "default=noprint_wrappers=1:nokey=1",

            self.media_path

        ]

        result = subprocess.run(
            command,
            capture_output=True,
            text=True
        )

        duration = float(result.stdout.strip())

        start = max(0, duration - seconds)

        return self.trim_audio(start, duration)

    # ---------------------------------------
    # Keep Last N Seconds (Video)
    # ---------------------------------------
    def keep_last_video(self, seconds):

        command = [

            "ffprobe",

            "-v", "error",

            "-show_entries",
            "format=duration",

            "-of",
            "default=noprint_wrappers=1:nokey=1",

            self.media_path

        ]

        result = subprocess.run(
            command,
            capture_output=True,
            text=True
        )

        duration = float(result.stdout.strip())

        start = max(0, duration - seconds)

        return self.trim_video(start, duration)