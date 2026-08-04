from django import forms
from .models import VideoAudio


class VideoAudioForm(forms.ModelForm):

    class Meta:

        model = VideoAudio

        fields = [
            "video",
            "audio",
            "mode",
            "volume",
            "loop_audio",
            
        ]

        widgets = {

            "video": forms.FileInput(
                attrs={
                    "class": "form-control",
                    "accept": "video/*"
                }
            ),

            "audio": forms.FileInput(
                attrs={
                    "class": "form-control",
                    "accept": "audio/*"
                }
            ),

            "mode": forms.Select(
                attrs={
                    "class": "form-select"
                }
            ),

            "volume": forms.NumberInput(
                attrs={
                    "class": "form-control",
                    "step": "0.1",
                    "min": "0",
                    "max": "2",
                    "value": "1"
                }
            ),

            "loop_audio": forms.CheckboxInput(
                attrs={
                    "class": "form-check-input"
                }
            ),

            "start_time": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "00:00:00"
                }
            ),

            "end_time": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "00:00:00"
                }
            )

        }

        labels = {

            "video": "Upload Video",

            "audio": "Upload Audio",

            "mode": "Audio Mode",

            "volume": "Volume",

            "loop_audio": "Loop Audio",

            "start_time": "Start Time",

            "end_time": "End Time"

        }

#============Add Te================


FONT_CHOICES = [
    ("Arial", "Arial"),
    ("Verdana", "Verdana"),
    ("Tahoma", "Tahoma"),
    ("Georgia", "Georgia"),
    ("Times New Roman", "Times New Roman"),
    ("Courier New", "Courier New"),
]


COLOR_CHOICES = [
    ("white", "White"),
    ("black", "Black"),
    ("red", "Red"),
    ("green", "Green"),
    ("blue", "Blue"),
    ("yellow", "Yellow"),
]


POSITION_CHOICES = [
    ("center", "Center"),
    ("top", "Top"),
    ("bottom", "Bottom"),
    ("left", "Left"),
    ("right", "Right"),
    ("top_left", "Top Left"),
    ("top_right", "Top Right"),
    ("bottom_left", "Bottom Left"),
    ("bottom_right", "Bottom Right"),
]


class AddTextVideoForm(forms.Form):

    video = forms.FileField(
        label="Video"
    )

    text = forms.CharField(
        max_length=200,
        widget=forms.TextInput(
            attrs={
                "placeholder": "Enter Text"
            }
        )
    )

   

    font_size = forms.IntegerField(
        min_value=10,
        max_value=100,
        initial=20
    )

    font_color = forms.ChoiceField(
        choices=COLOR_CHOICES,
        initial="white"
    )

    position = forms.ChoiceField(
        choices=POSITION_CHOICES,
        initial="center"
    )

    margin_x = forms.IntegerField(
        min_value=0,
        initial=20
    )

    margin_y = forms.IntegerField(
        min_value=0,
        initial=20
    )

    opacity = forms.IntegerField(
        min_value=0,
        max_value=100,
        initial=100
    )

    duration = forms.IntegerField(
        min_value=1,
        initial=10
    )

    def clean_video(self):

        video = self.cleaned_data["video"]

        allowed_extensions = [
            ".mp4",
            ".mov",
            ".avi",
            ".mkv",
            ".webm"
        ]

        filename = video.name.lower()

        if not any(filename.endswith(ext) for ext in allowed_extensions):
            raise forms.ValidationError(
                "Unsupported video format."
            )

        return video

    def clean_text(self):

        text = self.cleaned_data["text"].strip()

        if len(text) == 0:
            raise forms.ValidationError(
                "Text cannot be empty."
            )

        return text