from django import forms


class TrimVideoForm(forms.Form):
    """
    Form for trimming a video between a start
    and end time.
    """

    video = forms.FileField(
        label="Upload Video",
        required=True,
        widget=forms.ClearableFileInput(
            attrs={
                "accept": "video/*",
            }
        ),
    )

    start_time = forms.DecimalField(
        label="Start Time (seconds)",
        required=True,
        min_value=0,
        max_digits=12,
        decimal_places=3,
        widget=forms.NumberInput(
            attrs={
                "min": "0",
                "step": "0.001",
                "placeholder": "Example: 10",
            }
        ),
    )

    end_time = forms.DecimalField(
        label="End Time (seconds)",
        required=True,
        min_value=0,
        max_digits=12,
        decimal_places=3,
        widget=forms.NumberInput(
            attrs={
                "min": "0",
                "step": "0.001",
                "placeholder": "Example: 30",
            }
        ),
    )

    def clean_video(self):
        video = self.cleaned_data.get("video")

        if not video:
            raise forms.ValidationError(
                "Please upload a video file."
            )

        max_size = 500 * 1024 * 1024

        if video.size > max_size:
            raise forms.ValidationError(
                "Video file size must not exceed 500 MB."
            )

        allowed_extensions = {
            ".mp4",
            ".mov",
            ".avi",
            ".mkv",
            ".webm",
            ".flv",
            ".wmv",
            ".m4v",
        }

        file_name = video.name.lower()

        if "." not in file_name:
            raise forms.ValidationError(
                "The uploaded file has no valid extension."
            )

        extension = "." + file_name.rsplit(".", 1)[1]

        if extension not in allowed_extensions:
            raise forms.ValidationError(
                "Unsupported video format."
            )

        return video

    def clean(self):
        cleaned_data = super().clean()

        start_time = cleaned_data.get("start_time")
        end_time = cleaned_data.get("end_time")

        if (
            start_time is not None
            and end_time is not None
        ):

            if start_time >= end_time:

                raise forms.ValidationError(
                    "End time must be greater than start time."
                )

        return cleaned_data

from django import forms


class MultipleFileInput(forms.ClearableFileInput):
    allow_multiple_selected = True


class MultipleFileField(forms.FileField):
    widget = MultipleFileInput

    def clean(self, data, initial=None):
        single_file_clean = super().clean

        if not data:
            return []

        if isinstance(data, (list, tuple)):
            return [
                single_file_clean(file, initial)
                for file in data
            ]

        return [single_file_clean(data, initial)]



class MergeVideoForm(forms.Form):

    videos = MultipleFileField(
        required=True,
        widget=MultipleFileInput(
            attrs={
                "accept": "video/*",
            }
        ),
    )

    def clean_videos(self):

        videos = self.cleaned_data.get("videos")

        if not videos:
            raise forms.ValidationError(
                "Please upload at least two videos."
            )

        if len(videos) < 2:
            raise forms.ValidationError(
                "Please upload at least two videos."
            )

        return videos