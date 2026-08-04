import os
import subprocess
from django.conf import settings


class VolumeController:

    def __init__(self, audio_path):
        self.audio_path = audio_path

    # ---------------------------------------
    # Set Custom Volume
    # Example:
    # 1.0 = 100%
    # 0.5 = 50%
    # 2.0 = 200%
    # ---------------------------------------
    def set_volume(self, volume=1.0):

        output_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(output_dir, exist_ok=True)

        output_audio = os.path.join(
            output_dir,
            "volume_adjusted.wav"
        )

        command = [

            "ffmpeg",

            "-y",

            "-i",
            self.audio_path,

            "-filter:a",
            f"volume={volume}",

            output_audio

        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_audio

    # ---------------------------------------
    # Increase Volume
    # ---------------------------------------
    def increase(self, percent=150):

        volume = percent / 100

        return self.set_volume(volume)

    # ---------------------------------------
    # Decrease Volume
    # ---------------------------------------
    def decrease(self, percent=50):

        volume = percent / 100

        return self.set_volume(volume)

    # ---------------------------------------
    # Mute Audio
    # ---------------------------------------
    def mute(self):

        return self.set_volume(0)

    # ---------------------------------------
    # Double Volume
    # ---------------------------------------
    def double(self):

        return self.set_volume(2.0)

    # ---------------------------------------
    # Half Volume
    # ---------------------------------------
    def half(self):

        return self.set_volume(0.5)