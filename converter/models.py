from django.db import models
from django.core.validators import FileExtensionValidator

# This app uses file-based processing and doesn't require database models.
# Files are temporarily stored in MEDIA_ROOT during conversion.

class HeroVideo(models.Model):
    """
    Stores videos used in the homepage hero section playlist.
    """
    SECTION_CHOICES = [
        ('short_url', 'Short URL Hero'),
        ('qr_code', 'QR Code Hero'),
    ]
    section = models.CharField(
        max_length=20,
        choices=SECTION_CHOICES,
        default='short_url',
        help_text="Select where this video should be displayed."
    )
    title = models.CharField(max_length=200, blank=True)
    video = models.FileField(
        upload_to="hero_videos/",
        validators=[FileExtensionValidator(allowed_extensions=['mp4', 'webm', 'mov', 'm4v'])]
    )
    order = models.PositiveIntegerField(
        default=0,
        help_text="Lower numbers appear first in the playlist."
    )
    is_active = models.BooleanField(
        default=True,
        help_text="Uncheck to hide this video from the homepage."
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["order", "id"]
        verbose_name = "Hero Video"
        verbose_name_plural = "Hero Videos"

    def __str__(self):
        return self.title or str(self.video.name).split('/')[-1]
