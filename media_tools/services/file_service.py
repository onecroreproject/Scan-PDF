import logging
import uuid
from pathlib import Path

from django.conf import settings


logger = logging.getLogger("media_tools")


def get_video_directories():
    """
    Create and return video processing directories.
    """

    try:
        base_dir = (
            Path(settings.MEDIA_ROOT)
            / "video_tools"
        )

        uploads_dir = base_dir / "uploads"
        outputs_dir = base_dir / "outputs"
        temp_dir = base_dir / "temp"

        uploads_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        outputs_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        temp_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        return (
            uploads_dir,
            outputs_dir,
            temp_dir,
        )

    except OSError as exc:
        logger.exception(
            "Unable to create media directories."
        )
        raise RuntimeError(
            "Unable to prepare video storage."
        ) from exc

    except Exception as exc:
        logger.exception(
            "Unexpected file service error."
        )
        raise RuntimeError(
            "Unable to prepare video storage."
        ) from exc


def create_unique_filename(
    extension=".mp4",
):
    """
    Create a unique internal filename.
    """

    try:
        extension = extension.lower()

        if not extension.startswith("."):
            extension = f".{extension}"

        return (
            f"{uuid.uuid4().hex}"
            f"{extension}"
        )

    except Exception as exc:
        logger.exception(
            "Unable to create unique filename."
        )
        raise RuntimeError(
            "Unable to create output filename."
        ) from exc


def save_uploaded_file(uploaded_file):
    """
    Save uploaded file using a UUID filename.
    """

    try:
        uploads_dir, _, _ = (
            get_video_directories()
        )

        extension = (
            Path(uploaded_file.name)
            .suffix
            .lower()
        )

        filename = create_unique_filename(
            extension
        )

        file_path = uploads_dir / filename

        with file_path.open(
            "wb"
        ) as destination:

            for chunk in uploaded_file.chunks():
                destination.write(chunk)

        logger.info(
            "Uploaded video saved: %s",
            file_path,
        )

        return file_path

    except OSError as exc:
        logger.exception(
            "Unable to save uploaded video."
        )
        raise RuntimeError(
            "Unable to save uploaded video."
        ) from exc

    except Exception as exc:
        logger.exception(
            "Unexpected upload error."
        )
        raise RuntimeError(
            "Unable to save uploaded video."
        ) from exc