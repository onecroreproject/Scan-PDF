from django import forms
from django.conf import settings

from media_tools.utils.validators import (
    validate_video_file,
)


class BaseVideoForm(forms.Form):
    """
    Common form for all video tools.
    """

    video = forms.FileField(
        required=True,
        validators=[
            validate_video_file
        ],
    )

    def clean_video(self):
        try:
            video = self.cleaned_data["video"]

            max_size = getattr(
                settings,
                "MEDIA_TOOLS_MAX_FILE_SIZE",
                500 * 1024 * 1024,
            )

            if video.size > max_size:
                raise forms.ValidationError(
                    "Video file is too large."
                )

            return video

        except forms.ValidationError:
            raise

        except Exception as exc:
            raise forms.ValidationError(
                "Unable to validate uploaded video."
            ) from exc