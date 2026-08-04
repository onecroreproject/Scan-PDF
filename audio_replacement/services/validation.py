import os

from django.core.exceptions import ValidationError


class FileValidator:

    # Supported Formats
    VIDEO_EXTENSIONS = [
        ".mp4",
        ".avi",
        ".mov",
        ".mkv",
        ".webm"
    ]

    AUDIO_EXTENSIONS = [
        ".mp3",
        ".wav",
        ".aac",
        ".m4a",
        ".ogg",
        ".flac"
    ]

    # Maximum File Size (500 MB)
    MAX_VIDEO_SIZE = 500 * 1024 * 1024

    # Maximum Audio Size (100 MB)
    MAX_AUDIO_SIZE = 100 * 1024 * 1024

    # -----------------------------------------
    # Validate Video
    # -----------------------------------------
    @classmethod
    def validate_video(cls, video):

        if not video:
            raise ValidationError("Please upload a video.")

        extension = os.path.splitext(video.name)[1].lower()

        if extension not in cls.VIDEO_EXTENSIONS:
            raise ValidationError(
                "Unsupported video format."
            )

        if video.size > cls.MAX_VIDEO_SIZE:
            raise ValidationError(
                "Video size must be less than 500 MB."
            )

        return True

    # -----------------------------------------
    # Validate Audio
    # -----------------------------------------
    @classmethod
    def validate_audio(cls, audio):

        if not audio:
            raise ValidationError("Please upload an audio file.")

        extension = os.path.splitext(audio.name)[1].lower()

        if extension not in cls.AUDIO_EXTENSIONS:
            raise ValidationError(
                "Unsupported audio format."
            )

        if audio.size > cls.MAX_AUDIO_SIZE:
            raise ValidationError(
                "Audio size must be less than 100 MB."
            )

        return True

    # -----------------------------------------
    # Validate Both Files
    # -----------------------------------------
    @classmethod
    def validate(cls, video, audio):

        cls.validate_video(video)

        cls.validate_audio(audio)

        return True