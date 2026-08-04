import os
import cv2
import subprocess
from django.conf import settings


class VideoProcessor:

    def __init__(self, video_path):
        self.video_path = video_path

    # Check whether uploaded video exists
    def exists(self):
        return os.path.exists(self.video_path)

    # Read video information
    def get_video_info(self):

        if not self.exists():
            raise FileNotFoundError("Video file not found.")

        cap = cv2.VideoCapture(self.video_path)

        width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        fps = cap.get(cv2.CAP_PROP_FPS)
        frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))

        duration = frames / fps if fps else 0

        cap.release()

        return {
            "width": width,
            "height": height,
            "fps": round(fps, 2),
            "frames": frames,
            "duration": round(duration, 2)
        }

    # Extract original audio from uploaded video
    def extract_audio(self):

        temp_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(temp_dir, exist_ok=True)

        output_audio = os.path.join(temp_dir, "original_audio.wav")

        command = [
            "ffmpeg",
            "-y",
            "-i", self.video_path,
            "-vn",
            output_audio
        ]

        subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        return output_audio

    # Get only video duration
    def get_duration(self):
        return self.get_video_info()["duration"]

    # Get video resolution
    def get_resolution(self):
        info = self.get_video_info()
        return f'{info["width"]}x{info["height"]}'

    # Get FPS
    def get_fps(self):
        return self.get_video_info()["fps"]