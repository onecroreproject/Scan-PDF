import os
import subprocess
from django.conf import settings


class AudioLooper:

    def __init__(self, audio_path):
        self.audio_path = audio_path

    # ---------------------------------------
    # Get Audio Duration
    # ---------------------------------------
    def get_duration(self):

        command = [
            "ffprobe",
            "-v", "error",
            "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1",
            self.audio_path
        ]

        result = subprocess.run(
            command,
            capture_output=True,
            text=True
        )

        try:
            return float(result.stdout.strip())
        except:
            return 0

    # ---------------------------------------
    # Loop Audio for Specific Duration
    # Example:
    # video_duration = 120 (seconds)
    # ---------------------------------------
    def loop_to_duration(self, video_duration):

        output_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(output_dir, exist_ok=True)

        output_audio = os.path.join(
            output_dir,
            "looped_audio.wav"
        )

        command = [

            "ffmpeg",

            "-y",

            "-stream_loop",
            "-1",

            "-i",
            self.audio_path,

            "-t",
            str(video_duration),

            "-c",
            "copy",

            output_audio

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_audio

    # ---------------------------------------
    # Loop Audio N Times
    # ---------------------------------------
    def loop_times(self, count=2):

        output_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(output_dir, exist_ok=True)

        temp_list = os.path.join(
            output_dir,
            "audio_list.txt"
        )

        with open(temp_list, "w") as file:
            for _ in range(count):
                file.write(f"file '{self.audio_path}'\n")

        output_audio = os.path.join(
            output_dir,
            "loop_times.wav"
        )

        command = [

            "ffmpeg",

            "-y",

            "-f",
            "concat",

            "-safe",
            "0",

            "-i",
            temp_list,

            "-c",
            "copy",

            output_audio

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_audio

    # ---------------------------------------
    # Auto Loop
    # If audio is shorter than video,
    # repeat automatically.
    # ---------------------------------------
    def auto_loop(self, video_duration):

        audio_duration = self.get_duration()

        if audio_duration >= video_duration:
            return self.audio_path

        return self.loop_to_duration(video_duration)