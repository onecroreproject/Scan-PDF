"""
Django settings for allinone project.
"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = 'django-insecure-change-this-in-production-x9$k2m!q@w3e4r5t6y7u8i9o0p'

DEBUG = True

ALLOWED_HOSTS = ['*']

# CSRF security for production (Required for POST requests on HTTPS)
CSRF_TRUSTED_ORIGINS = [
    'https://scanpdf.co.in',
    'https://www.scanpdf.co.in',
    'http://scanpdf.co.in',
]

# Security settings
SECURE_PROXY_SSL_HEADER = ('HTTP_X_FORWARDED_PROTO', 'https')
CSRF_COOKIE_SECURE = not DEBUG
SESSION_COOKIE_SECURE = not DEBUG

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'converter',
    'image_processor',
    'audio_processor',
    'dynamic_qr',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',  # For production static files
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'allinone.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / 'templates'],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
                'converter.context_processors.tools_processor',
            ],
        },
    },
]

WSGI_APPLICATION = 'allinone.wsgi.application'

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
]

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'Asia/Kolkata'
USE_I18N = True
USE_TZ = True

# Static files
STATIC_URL = '/static/'
STATICFILES_DIRS = [BASE_DIR / 'static']
STATIC_ROOT = BASE_DIR / 'staticfiles'

# Media files (Redirected to system temp to keep project folder clean)
import tempfile
MEDIA_URL = '/media/'
MEDIA_ROOT = os.path.join(tempfile.gettempdir(), 'scanpdf_media_root')

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# Email configuration for Dynamic QR OTP (Gmail SMTP)
EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True
EMAIL_HOST_USER = os.environ.get('EMAIL_HOST_USER', '')      # your Gmail address
EMAIL_HOST_PASSWORD = os.environ.get('EMAIL_HOST_PASSWORD', '')  # Gmail app password
DEFAULT_FROM_EMAIL = os.environ.get('EMAIL_HOST_USER', 'noreply@scanpdf.co.in')

# File upload settings
FILE_UPLOAD_MAX_MEMORY_SIZE = 52428800  # 50 MB
DATA_UPLOAD_MAX_MEMORY_SIZE = 52428800  # 50 MB

# -----------------------------------------------------------------------------
# Local media tool binaries (FFmpeg)
# -----------------------------------------------------------------------------
import warnings


def _binary_name(base: str) -> str:
    return f"{base}.exe" if os.name == "nt" else base


# Preferred: bundle binaries inside the project:
# - Windows: ffmpeg/bin/ffmpeg.exe + ffmpeg/bin/ffprobe.exe
# - Linux:   ffmpeg/bin/ffmpeg     + ffmpeg/bin/ffprobe
FFMPEG_BIN_DIR = BASE_DIR / "ffmpeg" / "bin"
FFMPEG_PATH = str(FFMPEG_BIN_DIR / _binary_name("ffmpeg"))
FFPROBE_PATH = str(FFMPEG_BIN_DIR / _binary_name("ffprobe"))

# Optional: if you created ffmpeg/bin but forgot the binaries, fail loudly in DEBUG.
if FFMPEG_BIN_DIR.exists():
    if DEBUG:
        if not Path(FFMPEG_PATH).exists():
            raise RuntimeError(f"FFmpeg not found at {FFMPEG_PATH}")
    else:
        if not Path(FFMPEG_PATH).exists():
            warnings.warn(f"FFmpeg not found at {FFMPEG_PATH}; falling back to PATH/imageio-ffmpeg")
