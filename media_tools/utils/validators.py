from pathlib import Path

from django.core.exceptions import ValidationError


ALLOWED_VIDEO_EXTENSIONS = {
    ".mp4",
    ".mov",
    ".mkv",
    ".avi",
    ".webm",
    ".m4v",
}


def validate_video_file(uploaded_file):
    """
    Validate uploaded video file.
    """

    try:
        if not uploaded_file:
            raise ValidationError(
                "Please select a video file."
            )

        extension = Path(
            uploaded_file.name
        ).suffix.lower()

        if extension not in ALLOWED_VIDEO_EXTENSIONS:
            raise ValidationError(
                "Unsupported video format."
            )

        if uploaded_file.size <= 0:
            raise ValidationError(
                "The uploaded video is empty."
            )

    except ValidationError:
        raise

    except Exception as exc:
        raise ValidationError(
            "Unable to validate the video file."
        ) from exc