from django.db import models


class VideoAudio(models.Model):

    MODE_CHOICES = (
        ("replace", "Replace Audio"),
        ("merge", "Merge Audio"),
    )

    video = models.FileField(
        upload_to="videos/"
    )

    audio = models.FileField(
        upload_to="audios/"
    )

    output_video = models.FileField(
        upload_to="output/",
        blank=True,
        null=True
    )

    mode = models.CharField(
        max_length=20,
        choices=MODE_CHOICES,
        default="replace"
    )

    volume = models.FloatField(
        default=1.0
    )

    loop_audio = models.BooleanField(
        default=False
    )

    start_time = models.CharField(
        max_length=20,
        default="00:00:00"
    )

    end_time = models.CharField(
        max_length=20,
        default="00:00:00"
    )

    status = models.CharField(
        max_length=20,
        default="Pending"
    )

    created_at = models.DateTimeField(
        auto_now_add=True
    )

    updated_at = models.DateTimeField(
        auto_now=True
    )

    def __str__(self):
        return f"Video #{self.id}"