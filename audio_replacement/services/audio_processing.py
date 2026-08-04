import os
import subprocess
from django.conf import settings


class AudioProcessor:

    def __init__(self, audio_path):
        self.audio_path = audio_path

    # -------------------------------------------------
    # Check whether audio file exists
    # -------------------------------------------------
    def exists(self):
        return os.path.exists(self.audio_path)

    # -------------------------------------------------
    # Convert audio to WAV format
    # -------------------------------------------------
    def convert_to_wav(self):

        if not self.exists():
            raise FileNotFoundError("Audio file not found.")

        temp_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(temp_dir, exist_ok=True)

        output_audio = os.path.join(temp_dir, "processed_audio.wav")

        command = [
            "ffmpeg",
            "-y",
            "-i", self.audio_path,
            "-ac", "2",
            "-ar", "44100",
            output_audio
        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_audio

    # -------------------------------------------------
    # Get Audio Duration
    # -------------------------------------------------
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
            return round(float(result.stdout.strip()), 2)
        except:
            return 0

    # -------------------------------------------------
    # Normalize Audio Volume
    # -------------------------------------------------
    def normalize_audio(self):

        temp_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(temp_dir, exist_ok=True)

        output_audio = os.path.join(temp_dir, "normalized_audio.wav")

        command = [
            "ffmpeg",
            "-y",
            "-i", self.audio_path,
            "-filter:a", "loudnorm",
            output_audio
        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_audio

    # -------------------------------------------------
    # Remove Silence (Beginning & End)
    # -------------------------------------------------
    def remove_silence(self):

        temp_dir = os.path.join(settings.MEDIA_ROOT, "temp")
        os.makedirs(temp_dir, exist_ok=True)

        output_audio = os.path.join(temp_dir, "trimmed_audio.wav")

        command = [
            "ffmpeg",
            "-y",
            "-i", self.audio_path,
            "-af",
            "silenceremove=start_periods=1:start_duration=0.5:start_threshold=-40dB",
            output_audio
        ]

        subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        return output_audio

    # -------------------------------------------------
    # Complete Audio Processing Pipeline
    # -------------------------------------------------
    def process(self):

        wav_audio = self.convert_to_wav()

        processor = AudioProcessor(wav_audio)

        normalized = processor.normalize_audio()

        processor = AudioProcessor(normalized)

        final_audio = processor.remove_silence()

        return final_audio