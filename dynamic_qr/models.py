import uuid
import string
import random
from django.db import models
from django.contrib.auth.models import User


def generate_short_code():
    """Generate a random 8-character short code for QR redirect URLs."""
    chars = string.ascii_letters + string.digits
    return ''.join(random.choices(chars, k=8))


class DynamicQRCode(models.Model):
    """
    Stores dynamic QR code data.
    Only dynamic QR data is stored in the database — no other project data.
    """
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='dynamic_qr_codes')
    short_code = models.CharField(max_length=20, unique=True, default=generate_short_code, db_index=True)

    # QR content
    qr_name = models.CharField(max_length=200, help_text="A friendly name for this QR code")
    qr_type = models.CharField(max_length=30, default='url', choices=[
        ('url', 'URL'),
        ('text', 'Text'),
        ('email', 'Email'),
        ('phone', 'Phone'),
        ('sms', 'SMS'),
        ('wifi', 'WiFi'),
        ('vcard', 'vCard'),
        ('location', 'Location'),
    ])
    destination_url = models.URLField(max_length=2000, blank=True, null=True, help_text="Primary URL for direct redirect types")
    qr_data = models.JSONField(default=dict, blank=True, help_text="Structured data for non-URL QR types")

    # Design options stored as JSON
    fg_color = models.CharField(max_length=10, default='#000000')
    bg_color = models.CharField(max_length=10, default='#ffffff')
    body_style = models.CharField(max_length=20, default='square')
    eye_style = models.CharField(max_length=20, default='square')
    ball_style = models.CharField(max_length=20, default='square')
    logo = models.ImageField(upload_to='dynamic_qr_logos/', null=True, blank=True)

    # Analytics
    scan_count = models.PositiveIntegerField(default=0)

    # Status
    is_active = models.BooleanField(default=True)

    # Timestamps
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Dynamic QR Code'
        verbose_name_plural = 'Dynamic QR Codes'

    def __str__(self):
        return f"{self.qr_name} ({self.short_code})"

    def increment_scan(self):
        self.scan_count += 1
        self.save(update_fields=['scan_count'])


class OTPVerification(models.Model):
    """
    Stores OTP codes for forgot-password email verification.
    Only used for the dynamic QR feature's auth system.
    """
    email = models.EmailField()
    otp_code = models.CharField(max_length=6)
    created_at = models.DateTimeField(auto_now_add=True)
    is_used = models.BooleanField(default=False)
    attempts = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'OTP Verification'

    def __str__(self):
        return f"OTP for {self.email}"
