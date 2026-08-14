"""
Django settings for allinone project.
"""
import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(os.path.join(BASE_DIR, '.env'))

SECRET_KEY = 'django-insecure-change-this-in-production-x9$k2m!q@w3e4r5t6y7u8i9o0p'

DEBUG = os.environ.get('DEBUG', 'True').lower() == 'true'

ALLOWED_HOSTS = ['*']

# CSRF security for production (Required for POST requests on HTTPS)


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
    'video_downloader',

    'services',
    'custom_admin',

    'audio_replacement',
    'media_tools',
    #"media_processing.apps.MediaProcessingConfig",

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
    'services.middleware.SubscriptionMiddleware',
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

# Static files (CSS, JavaScript, Images)
STATIC_URL = '/static/'
STATICFILES_DIRS = [BASE_DIR / 'static']
STATIC_ROOT = BASE_DIR / 'staticfiles'

# Static files storage
if DEBUG:
    STATICFILES_STORAGE = 'django.contrib.staticfiles.storage.StaticFilesStorage'
else:
    STATICFILES_STORAGE = 'whitenoise.storage.CompressedManifestStaticFilesStorage'

# WhiteNoise Configuration for maximum reliability
WHITENOISE_MANIFEST_STRICT = False
WHITENOISE_USE_FINDERS = True
WHITENOISE_AUTOREFRESH = DEBUG

# Ensure correct MIME types on all systems (especially Windows)
import mimetypes
mimetypes.add_type("text/css", ".css", True)
mimetypes.add_type("application/javascript", ".js", True)

# Media files (Redirected to system temp to keep project folder clean)
MEDIA_URL = "/media/"
MEDIA_ROOT = os.path.join(BASE_DIR, "media")

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# Email configuration for Dynamic QR OTP (Gmail SMTP)
EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True
EMAIL_HOST_USER = os.environ.get('EMAIL_HOST_USER', '')      # your Gmail address
EMAIL_HOST_PASSWORD = os.environ.get('EMAIL_HOST_PASSWORD', '').replace(' ', '')  # Gmail app password (no spaces)
DEFAULT_FROM_EMAIL = EMAIL_HOST_USER

# File upload settings
FILE_UPLOAD_MAX_MEMORY_SIZE = 524288000  # 500 MB - store large uploads on disk
DATA_UPLOAD_MAX_MEMORY_SIZE = 524288000  # 500 MB
#FILE_UPLOAD_TEMP_DIR = os.path.join(tempfile.gettempdir(), 'scanpdf_uploads')
#os.makedirs(FILE_UPLOAD_TEMP_DIR, exist_ok=True)

DATA_UPLOAD_MAX_NUMBER_FIELDS = 10000

# ═══════════════════════════════════════════════════════════════
# CELERY CONFIGURATION
# ═══════════════════════════════════════════════════════════════
CELERY_BROKER_URL = os.environ.get('CELERY_BROKER_URL', 'redis://localhost:6379/0')
CELERY_RESULT_BACKEND = os.environ.get('CELERY_RESULT_BACKEND', 'redis://localhost:6379/0')
CELERY_ACCEPT_CONTENT = ['json']
CELERY_TASK_SERIALIZER = 'json'
CELERY_RESULT_SERIALIZER = 'json'
CELERY_TASK_TRACK_STARTED = True
CELERY_TASK_TIME_LIMIT = 3600  # 1 hour max per task
CELERY_WORKER_PREFETCH_MULTIPLIER = 1  # For large file tasks, don't prefetch
CELERY_BROKER_CONNECTION_RETRY_ON_STARTUP = True
VIDEO_TEMP_MAX_AGE = 600  # 10 minutes

# ═══════════════════════════════════════════════════════════════
# LOGGING
# ═══════════════════════════════════════════════════════════════
LOGGING = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'verbose': {
            'format': '{levelname} {asctime} {module} {message}',
            'style': '{',
        },
    },
    'handlers': {
        'console': {
            'level': 'INFO',
            'class': 'logging.StreamHandler',
            'formatter': 'verbose',
        },
    },
    'loggers': {},
}

# Ensure logs directory exists
os.makedirs(BASE_DIR / 'logs', exist_ok=True)

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
