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
    QR_TYPES = [
        # Basic
        ('url', 'URL / Link'),
        ('text', 'Text Content'),
        ('wifi', 'Wi-Fi Network'),
        ('location', 'Map / Location'),
        
        # Social & Media
        ('whatsapp', 'WhatsApp'),
        ('youtube', 'YouTube'),
        ('facebook', 'Facebook'),
        ('instagram', 'Instagram'),
        ('telegram', 'Telegram'),
        ('tiktok', 'TikTok'),
        ('x-twitter', 'X / Twitter'),
        ('snapchat', 'Snapchat'),
        ('pinterest', 'Pinterest'),
        ('linkedin', 'LinkedIn'),
        
        # Files (Uploads)
        ('pdf', 'PDF Document'),
        ('audio', 'Audio / MP3'),
        ('video', 'Video / MP4'),
        ('image', 'Image / Photo'),
        ('pptx', 'PowerPoint (PPTX)'),
        ('excel', 'Excel (XLSX)'),
        ('word', 'Word (DOCX)'),
        
        # Communication & Utility
        ('email', 'Email Address'),
        ('phone', 'Phone Call'),
        ('sms', 'SMS Message'),
        ('vcard', 'vCard (Contact)'),
        ('calendar', 'Calendar Event'),
        ('booking', 'Booking / Appt'),
        
        # Business & Payments
        ('google-review', 'Google Review'),
        ('google-forms', 'Google Forms'),
        ('google-doc', 'Google Doc'),
        ('google-sheets', 'Google Sheets'),
        ('play-market', 'Play Market'),
        ('app-store', 'App Store'),
        ('paypal', 'PayPal'),
        ('etsy', 'Etsy Shop'),
        ('amazon', 'Amazon Product'),
        ('venmo', 'Venmo'),
        ('upi', 'UPI Payment'),
        ('crypto', 'Crypto Payment'),
        ('spotify', 'Spotify'),
        
        # Advanced
        ('link-list', 'List of Links'),
        ('custom-url', 'Custom Short URL'),
        ('office-365', 'Office 365'),
        ('2d-barcode', '2D Barcode'),
    ]
    qr_type = models.CharField(max_length=40, default='url', choices=QR_TYPES)
    destination_url = models.URLField(max_length=2000, blank=True, null=True, help_text="Primary URL for direct redirect types")
    qr_data = models.JSONField(default=dict, blank=True, help_text="Structured data for non-URL QR types")
    file_content = models.FileField(upload_to='dynamic_qr_contents/', null=True, blank=True, help_text="File associated with PDF, Audio, Video, etc.")

    # Design options stored as JSON
    fg_color = models.CharField(max_length=10, default='#000000')
    bg_color = models.CharField(max_length=10, default='#ffffff')
    body_style = models.CharField(max_length=20, default='square')
    eye_style = models.CharField(max_length=20, default='square')
    ball_style = models.CharField(max_length=20, default='square')
    logo = models.ImageField(upload_to='dynamic_qr_logos/', null=True, blank=True)
    design_options = models.JSONField(default=dict, blank=True, help_text="Advanced design options like frames, text, error correction, etc.")


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

    def get_static_content(self, request=None):
        """
        Determines the content to be encoded in the QR code.
        Returns the direct data for hardware-bound types (Wifi, vCard) 
        and the updateable redirect URL for content-bound types (Text, Web, Files).
        """
        # 1. Hardware-Bound Types (Must be static to trigger phone features)
        data = self.qr_data or {}
        if self.qr_type == 'wifi':
            ssid = data.get('ssid', '')
            pw = data.get('password', '')
            enc = data.get('encryption', 'WPA')
            return f"WIFI:S:{ssid};T:{enc};P:{pw};;"
        if self.qr_type == 'location':
            lat = data.get('latitude', '0')
            lon = data.get('longitude', '0')
            return f"geo:{lat},{lon}"
        if self.qr_type == 'vcard' and data.get('first_name'):
            fn = data.get('first_name')
            ln = data.get('last_name', '')
            org = data.get('organization', '')
            tel = data.get('phone_mobile', '')
            email = data.get('email', '')
            return f"BEGIN:VCARD\nVERSION:3.0\nN:{ln};{fn};;;\nFN:{fn} {ln}\nORG:{org}\nTEL;TYPE=CELL:{tel}\nEMAIL:{email}\nEND:VCARD"

        # 2. Content-Bound Types (Use redirect URL to allow updates after printing)
        if request:
            return request.build_absolute_uri(f'/qr/r/{self.short_code}/')
        return f"/qr/r/{self.short_code}/"


class QRAnalytics(models.Model):
    """
    Tracks individual scan events for dynamic QR codes.
    Stores metadata like IP address, user agent, and device type.
    """
    qr_code = models.ForeignKey(DynamicQRCode, on_delete=models.CASCADE, related_name='analytics')
    timestamp = models.DateTimeField(auto_now_add=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    user_agent = models.TextField(null=True, blank=True)
    browser = models.CharField(max_length=50, null=True, blank=True)
    os = models.CharField(max_length=50, null=True, blank=True)
    device_type = models.CharField(max_length=50, null=True, blank=True, help_text="Mobile, Tablet, Desktop")
    country = models.CharField(max_length=100, default='Unknown')
    city = models.CharField(max_length=100, default='Unknown')

    class Meta:
        ordering = ['-timestamp']
        verbose_name = 'QR Analytics'
        verbose_name_plural = 'QR Analytics'

    def __str__(self):
        return f"Scan for {self.qr_code.qr_name} at {self.timestamp}"


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
