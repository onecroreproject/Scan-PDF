import os
import time
import uuid
import logging
import subprocess
import json
import shutil
import platform
import traceback
from django.conf import settings
from urllib.parse import urlparse

try:
    import yt_dlp
except ImportError:
    yt_dlp = None

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Environment diagnostic helpers (safe – never log secret values)
# ---------------------------------------------------------------------------

def _get_yt_dlp_version():
    if yt_dlp is None:
        return 'NOT_INSTALLED'
    try:
        return yt_dlp.version.__version__
    except AttributeError:
        return getattr(yt_dlp, '__version__', 'unknown')


def _find_binary(name):
    """Return path if binary is found in PATH or FFMPEG_BIN_DIR, else None."""
    path = shutil.which(name)
    if path:
        return path
    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
        path = shutil.which(name, path=str(ffmpeg_dir))
    return path


def _deno_path():
    """Return Deno path if available, else None."""
    explicit = os.environ.get('DENO_PATH') or os.environ.get('DENO_INSTALL')
    if explicit:
        candidate = os.path.join(explicit, 'bin', 'deno') if not os.path.isfile(explicit) else explicit
        if os.path.isfile(candidate):
            return candidate
    return shutil.which('deno')


def _check_pot_provider():
    """Check if a PO Token provider plugin is discoverable. Safe – no secrets."""
    try:
        import importlib.util
        for mod in ('bgutil_ytdlp_pot_provider', 'yt_dlp_get_pot_rustypipe', 'yt_dlp_get_pot'):
            if importlib.util.find_spec(mod):
                return mod
        return None
    except Exception:
        return None


def log_environment_snapshot(url=None):
    """Safe diagnostic log – NEVER logs cookies, PO tokens, or secret values."""
    logger.info(
        "yt-dlp environment snapshot",
        extra={
            "python_executable": os.sys.executable,
            "python_version": platform.python_version(),
            "yt_dlp_version": _get_yt_dlp_version(),
            "os_platform": platform.platform(),
            "ffmpeg_path": _find_binary('ffmpeg') or "NOT_FOUND",
            "ffprobe_path": _find_binary('ffprobe') or "NOT_FOUND",
            "deno_path": _deno_path() or "NOT_FOUND",
            "pot_provider_detected": _check_pot_provider(),
            "target_url_domain": urlparse(url).netloc if url else "n/a",
            "cookie_file_configured": bool(
                os.environ.get('YTDLP_YOUTUBE_COOKIE_FILE') or
                os.environ.get('YTDLP_COOKIE_FILE')
            ),
        }
    )


# ---------------------------------------------------------------------------
# YTDLPError
# ---------------------------------------------------------------------------

class YTDLPError(Exception):
    def __init__(self, code, message, raw_error=None):
        self.code = code
        self.message = message
        self.raw_error = raw_error
        super().__init__(self.message)


# ---------------------------------------------------------------------------
# Error classification – deterministic, granular
# ---------------------------------------------------------------------------

_ERROR_MESSAGES = {
    'INVALID_URL':           'The URL provided is not a valid video link.',
    'UNSUPPORTED_URL':       'This URL is from an unsupported platform.',
    'VIDEO_UNAVAILABLE':     'This video is unavailable.',
    'PRIVATE_CONTENT':       'This video is private and cannot be accessed.',
    'AUTH_REQUIRED':         'This content requires authorized YouTube access.',
    'AGE_RESTRICTED':        'This content is age-restricted and cannot be accessed without account verification.',
    'GEO_RESTRICTED':        "This video is not available from this server's region.",
    'BOT_CHALLENGE':         'YouTube is requesting bot verification from this server. Please try again later.',
    'PO_TOKEN_REQUIRED':     "The server could not complete YouTube's current playback verification. Please try again.",
    'JS_CHALLENGE_FAILED':   "The server could not complete YouTube's current playback verification. Please try again.",
    'RATE_LIMITED':          'YouTube is temporarily limiting requests from this server. Please try again later.',
    'HTTP_403':              "The platform rejected the server's request (HTTP 403). Please try again later.",
    'HTTP_429':              'YouTube is temporarily limiting requests from this server. Please try again later.',
    'FORMAT_UNAVAILABLE':    'The requested format is not available for this video.',
    'FFMPEG_MISSING':        'A required media processing component (FFmpeg) is missing on the server.',
    'NETWORK_ERROR':         'The server could not reach the video platform. Please try again.',
    'DOWNLOAD_FAILED':       'The file download failed. Please try again.',
    'SERVER_CONFIGURATION':  'The server has a configuration issue. Please contact support.',
    'EXTRACTION_FAILED':     'Unable to extract video information. Please try again.',
    'METADATA_PARSE_FAILED': 'Unable to parse video metadata. Please try again.',
    'FFMPEG_REQUIRED':       'This quality requires media merging, which is currently unavailable on the server.',
}


def _categorize_error(e, url=''):
    """
    Classify a yt-dlp exception into a deterministic error code.

    CRITICALLY: 'sign in to confirm you are not a bot' is BOT_CHALLENGE,
    NOT AUTH_REQUIRED. These must be checked before generic 'sign in' patterns.
    """
    error_msg = str(e).lower()
    netloc = urlparse(url).netloc.lower() if url else ''

    # Facebook / Instagram specific
    if 'facebook.com' in netloc or 'instagram.com' in netloc:
        if any(k in error_msg for k in [
            'could not copy chrome cookie database', 'cookie database',
            'permission denied', 'cannot parse data',
        ]):
            return 'SERVER_CONFIGURATION', _ERROR_MESSAGES['SERVER_CONFIGURATION']
        if any(k in error_msg for k in ['private', 'login', 'sign in', 'authentication']):
            return 'AUTH_REQUIRED', _ERROR_MESSAGES['AUTH_REQUIRED']

    # Null / configuration
    if 'nonetype' in error_msg and 'youtubedl' in error_msg:
        return 'SERVER_CONFIGURATION', _ERROR_MESSAGES['SERVER_CONFIGURATION']

    # BOT_CHALLENGE must be checked BEFORE generic 'sign in'/'login' patterns.
    # YouTube returns "Sign in to confirm you're not a bot" for server IPs —
    # this is bot detection, NOT a requirement that the VIDEO needs a login.
    if any(k in error_msg for k in [
        "confirm you're not a bot",
        "confirm you are not a bot",
        'botguard',
        'sign in to confirm',
    ]):
        return 'BOT_CHALLENGE', _ERROR_MESSAGES['BOT_CHALLENGE']

    # PO Token
    if any(k in error_msg for k in ['po token', 'potoken', 'po_token']):
        return 'PO_TOKEN_REQUIRED', _ERROR_MESSAGES['PO_TOKEN_REQUIRED']

    # JS challenge
    if any(k in error_msg for k in ['js challenge', 'challenge failed', 'ejs', 'jschallenge']):
        return 'JS_CHALLENGE_FAILED', _ERROR_MESSAGES['JS_CHALLENGE_FAILED']

    # HTTP 429
    if 'http error 429' in error_msg or 'too many requests' in error_msg:
        return 'HTTP_429', _ERROR_MESSAGES['HTTP_429']

    # HTTP 400
    if 'http error 400' in error_msg or 'bad request' in error_msg:
        return 'NETWORK_ERROR', _ERROR_MESSAGES['NETWORK_ERROR']

    # HTTP 403
    if 'http error 403' in error_msg or ('forbidden' in error_msg and 'http' in error_msg):
        if any(k in error_msg for k in ['rate', 'captcha', 'bot', 'verify']):
            return 'BOT_CHALLENGE', _ERROR_MESSAGES['BOT_CHALLENGE']
        return 'HTTP_403', _ERROR_MESSAGES['HTTP_403']

    # Private
    if any(k in error_msg for k in ['private video', 'private content', 'members only']):
        return 'PRIVATE_CONTENT', _ERROR_MESSAGES['PRIVATE_CONTENT']

    # Age restricted
    if 'age' in error_msg and any(k in error_msg for k in ['restrict', 'gate', 'gated']):
        return 'AGE_RESTRICTED', _ERROR_MESSAGES['AGE_RESTRICTED']

    # Real auth (login for private/members content) – only after bot checks above
    if any(k in error_msg for k in ['login required', 'sign in to watch', 'logged-in']):
        return 'AUTH_REQUIRED', _ERROR_MESSAGES['AUTH_REQUIRED']

    # Generic bot/rate
    if any(k in error_msg for k in ['bot', 'verify', 'empty media response']):
        return 'BOT_CHALLENGE', _ERROR_MESSAGES['BOT_CHALLENGE']

    # Video unavailable
    if any(k in error_msg for k in [
        'video unavailable', 'unavailable video', 'not available',
        'no video formats', 'this video has been removed', 'deleted',
    ]):
        return 'VIDEO_UNAVAILABLE', _ERROR_MESSAGES['VIDEO_UNAVAILABLE']

    # Network
    if any(k in error_msg for k in ['network', 'timeout', 'timed out', 'connection']):
        return 'NETWORK_ERROR', _ERROR_MESSAGES['NETWORK_ERROR']

    # Geo restricted
    if any(k in error_msg for k in ['geo', 'country', 'region', 'not available in your country']):
        return 'GEO_RESTRICTED', _ERROR_MESSAGES['GEO_RESTRICTED']

    # Format unavailable
    if any(k in error_msg for k in ['format is not available', 'requested format']):
        return 'FORMAT_UNAVAILABLE', _ERROR_MESSAGES['FORMAT_UNAVAILABLE']

    # FFmpeg missing
    if any(k in error_msg for k in ['ffmpeg is not installed', 'ffmpeg not found', 'ffprobe and ffmpeg']):
        return 'FFMPEG_MISSING', _ERROR_MESSAGES['FFMPEG_MISSING']

    return 'EXTRACTION_FAILED', _ERROR_MESSAGES['EXTRACTION_FAILED']


# ---------------------------------------------------------------------------
# Base yt-dlp options
# ---------------------------------------------------------------------------

def get_ytdl_base_options(url):
    """
    Build yt-dlp options for the given URL.

    YouTube strategy (current 2025/2026 yt-dlp guidance):
      - Do NOT force random client rotation (tv, android, ios, web).
        This is fragile and hides the real issue.
      - Let yt-dlp choose its own default client.
      - If YTDLP_YOUTUBE_CLIENT env var is set, honour it (e.g. 'mweb').
      - If a PO Token Provider plugin is installed, yt-dlp auto-uses it.
      - Use a secure Netscape-format cookie FILE only if explicitly configured
        via environment variable. NEVER use cookiesfrombrowser on production
        (no Chrome/Firefox on Linux server).
    """
    options = {
        'quiet': True,
        'no_warnings': True,
        'nocheckcertificate': True,
        'geo_bypass': True,
        'retries': 3,
        'fragment_retries': 5,
        'socket_timeout': 30,
        'format_sort': ['vcodec:h264', 'res', 'acodec:m4a'],
    }

    parsed = urlparse(url)
    hostname = parsed.hostname.lower() if parsed.hostname else ''

    if 'youtube.com' in hostname or 'youtu.be' in hostname:
        # Only set extractor_args if an explicit client override is requested.
        # Default: let yt-dlp pick the best client for the installed version.
        explicit_client = os.environ.get('YTDLP_YOUTUBE_CLIENT', '').strip()
        if explicit_client:
            options['extractor_args'] = {
                'youtube': {'player_client': [explicit_client]}
            }

        cookies_file = (
            os.environ.get('YTDLP_YOUTUBE_COOKIE_FILE') or
            os.environ.get('YTDLP_COOKIE_FILE')
        )
        if cookies_file and os.path.isfile(cookies_file):
            options['cookiefile'] = cookies_file
            logger.info("YouTube cookie file applied.")

    elif 'facebook.com' in hostname or 'fb.watch' in hostname:
        cookies_file = (
            os.environ.get('YTDLP_FACEBOOK_COOKIE_FILE') or
            os.environ.get('YTDLP_COOKIE_FILE')
        )
        if cookies_file and os.path.isfile(cookies_file):
            options['cookiefile'] = cookies_file

    elif 'instagram.com' in hostname:
        cookies_file = (
            os.environ.get('YTDLP_INSTAGRAM_COOKIE_FILE') or
            os.environ.get('YTDLP_COOKIE_FILE')
        )
        if cookies_file and os.path.isfile(cookies_file):
            options['cookiefile'] = cookies_file

    elif 'twitter.com' in hostname or 'x.com' in hostname:
        cookies_file = (
            os.environ.get('YTDLP_TWITTER_COOKIE_FILE') or
            os.environ.get('YTDLP_COOKIE_FILE')
        )
        if cookies_file and os.path.isfile(cookies_file):
            options['cookiefile'] = cookies_file

    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
        options['ffmpeg_location'] = str(ffmpeg_dir)

    return options


# ---------------------------------------------------------------------------
# Core execution wrapper
# ---------------------------------------------------------------------------

def _execute_with_retry(execute_func, url, options):
    """Run the yt-dlp execute function; convert exceptions into YTDLPError."""
    try:
        return execute_func(options)
    except Exception as e:
        code, safe_msg = _categorize_error(e, url)
        raw_err_str = str(e) + "\n\n" + traceback.format_exc()

        # Safe diagnostic log – NO secret values logged
        logger.error(
            "yt-dlp operation failed",
            extra={
                "video_url_domain": urlparse(url).netloc,
                "error_code": code,
                "yt_dlp_version": _get_yt_dlp_version(),
                "ffmpeg_available": bool(_find_binary('ffmpeg')),
                "ffprobe_available": bool(_find_binary('ffprobe')),
                "deno_available": bool(_deno_path()),
                "os_platform": platform.platform(),
                "pot_provider": _check_pot_provider(),
                "error_class": type(e).__name__,
                "error_summary": str(e)[:200],
            }
        )
        raise YTDLPError(code, safe_msg, raw_error=raw_err_str)


# ---------------------------------------------------------------------------
# FFprobe helpers
# ---------------------------------------------------------------------------

def verify_video_audio_streams(filepath):
    """Use ffprobe to verify that the file contains the expected media streams."""
    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    ffprobe_cmd = _find_binary('ffprobe') or 'ffprobe'
    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
        candidate = os.path.join(str(ffmpeg_dir), 'ffprobe')
        if os.path.isfile(candidate):
            ffprobe_cmd = candidate

    cmd = [ffprobe_cmd, '-v', 'quiet', '-print_format', 'json', '-show_streams', filepath]

    try:
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode != 0:
            logger.error("ffprobe failed for %s: %s", filepath, result.stderr[:500])
            return False, False, None, None

        data = json.loads(result.stdout)
        streams = data.get('streams', [])

        has_video = False
        has_audio = False
        audio_codec = None
        video_codec = None

        for s in streams:
            codec_type = s.get('codec_type')
            if codec_type == 'video':
                has_video = True
                video_codec = s.get('codec_name', '').lower()
            elif codec_type == 'audio':
                has_audio = True
                audio_codec = s.get('codec_name', '').lower()

        return has_video, has_audio, audio_codec, video_codec
    except Exception as e:
        logger.error("Error running ffprobe on %s: %s", filepath, e)
        return False, False, None, None


def get_codec_priority(vcodec):
    if not vcodec or vcodec == 'none':
        return 0
    vcodec = vcodec.lower()
    if 'avc' in vcodec or 'h264' in vcodec:
        return 4
    if 'hev' in vcodec or 'h265' in vcodec:
        return 3
    if 'vp9' in vcodec:
        return 2
    if 'av01' in vcodec:
        return 1
    return 0


# ---------------------------------------------------------------------------
# Analyze
# ---------------------------------------------------------------------------

def analyze_video(url):
    """
    Extract metadata and available formats using yt-dlp (download=False).
    Returns a dict suitable for JSON serialisation.
    """
    if yt_dlp is None:
        raise YTDLPError('SERVER_CONFIGURATION', _ERROR_MESSAGES['SERVER_CONFIGURATION'])

    log_environment_snapshot(url)
    options = get_ytdl_base_options(url)

    def _extract(opts):
        with yt_dlp.YoutubeDL(opts) as ydl:
            return ydl.extract_info(url, download=False)

    try:
        info = _execute_with_retry(_extract, url, options)

        duration = info.get('duration') or 0
        if not duration:
            for f in info.get('formats', []):
                if f.get('duration'):
                    duration = f['duration']
                    break

        result = {
            'title': info.get('title', 'Unknown Title'),
            'thumbnail': info.get('thumbnail'),
            'duration': duration,
            'uploader': info.get('uploader', info.get('extractor_key')),
            'formats': [],
        }

        video_formats_by_height = {}
        audio_formats = []

        for f in info.get('formats', []):
            vcodec = str(f.get('vcodec') or 'none').lower()
            acodec = str(f.get('acodec') or 'none').lower()
            ext    = str(f.get('ext')    or 'unknown').lower()

            filesize = f.get('filesize') or f.get('filesize_approx') or 0
            height   = f.get('height')   or 0
            width    = f.get('width')    or 0
            bitrate  = f.get('tbr') or f.get('vbr') or f.get('abr') or 0

            # Some social extractors don't label codecs but output valid MP4s
            if vcodec == 'none' and acodec == 'none':
                if (height > 0 or width > 0) and ext in ['mp4', 'webm', 'mov']:
                    vcodec = 'unknown'
                else:
                    continue

            if vcodec == 'none' and acodec != 'none':
                abr = f.get('abr') or bitrate or 0
                approx_bitrate = round(abr / 32) * 32 if abr else 0
                audio_formats.append({
                    'format_id':  f.get('format_id'),
                    'resolution': f.get('format_note', 'Audio'),
                    'ext':        'mp3',
                    'vcodec':     'none',
                    'acodec':     'MP3',
                    'bitrate':    bitrate,
                    'filesize':   filesize,
                    'abr':        approx_bitrate,
                    'type':       'Audio Only',
                })
                continue

            if height > 0:
                priority = get_codec_priority(vcodec)
                display_height = width if (height > width and width > 0) else height

                fid = f.get('format_id')
                if acodec == 'none':
                    fid = f"{fid}+bestaudio/best"

                fmt = {
                    'format_id':  fid,
                    'resolution': f"{display_height}p",
                    'ext':        'mp4',
                    'vcodec':     'H.264' if priority >= 3 else vcodec,
                    'acodec':     'AAC',
                    'bitrate':    bitrate,
                    'filesize':   filesize,
                    'priority':   priority,
                    'type':       'Video + Audio',
                    'raw_acodec': acodec,
                }

                if display_height not in video_formats_by_height:
                    video_formats_by_height[display_height] = fmt
                elif priority > video_formats_by_height[display_height]['priority']:
                    video_formats_by_height[display_height] = fmt

        final_video_formats = [
            video_formats_by_height[h]
            for h in sorted(video_formats_by_height.keys(), reverse=True)
        ]

        has_any_audio = any(
            (f.get('acodec') and f.get('acodec') != 'none') or
            f.get('asr') or f.get('audio_channels')
            for f in info.get('formats', [])
        )
        if not has_any_audio and any(d in url.lower() for d in ['facebook', 'instagram']):
            has_any_audio = True

        seen_abr = set()
        final_audio_formats = []
        for fmt in sorted(audio_formats, key=lambda x: x['abr'], reverse=True):
            if fmt['abr'] not in seen_abr:
                seen_abr.add(fmt['abr'])
                final_audio_formats.append(fmt)

        if not final_audio_formats and has_any_audio:
            final_audio_formats.append({
                'format_id':  'bestaudio/best',
                'resolution': 'Best Audio',
                'ext':        'mp3',
                'vcodec':     'none',
                'acodec':     'MP3',
                'bitrate':    192,
                'filesize':   0,
                'type':       'Audio Only',
            })

        if not final_video_formats:
            final_video_formats.append({
                'format_id':  'bestvideo+bestaudio/best',
                'resolution': 'Highest Quality',
                'ext':        'mp4',
                'vcodec':     'H.264',
                'acodec':     'AAC',
                'bitrate':    0,
                'filesize':   0,
                'type':       'Video + Audio',
            })

        result['formats'] = final_video_formats + final_audio_formats
        return result

    except YTDLPError:
        raise
    except Exception as e:
        raw_err_str = str(e) + "\n\n" + traceback.format_exc()
        logger.error("Failed to analyze URL %s: %s", urlparse(url).netloc, str(e)[:200])
        raise YTDLPError('METADATA_PARSE_FAILED', _ERROR_MESSAGES['METADATA_PARSE_FAILED'], raw_error=raw_err_str)


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

def download_format(url, format_id, format_type, download_id=None):
    """
    Download the specified format.
    Returns (absolute_filepath, title).
    """
    if yt_dlp is None:
        raise YTDLPError('SERVER_CONFIGURATION', _ERROR_MESSAGES['SERVER_CONFIGURATION'])

    from django.core.cache import cache

    ffmpeg_available = bool(_find_binary('ffmpeg'))
    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    if not ffmpeg_available and ffmpeg_dir and os.path.exists(ffmpeg_dir):
        ffmpeg_available = bool(shutil.which('ffmpeg', path=str(ffmpeg_dir)))

    if format_type == 'Video + Audio' and '+bestaudio' in format_id and not ffmpeg_available:
        raise YTDLPError('FFMPEG_REQUIRED', _ERROR_MESSAGES['FFMPEG_REQUIRED'])

    if format_type == 'Audio Only' and not ffmpeg_available:
        raise YTDLPError('FFMPEG_REQUIRED', _ERROR_MESSAGES['FFMPEG_REQUIRED'])

    temp_dir = os.path.join(settings.MEDIA_ROOT, 'video_downloads')
    os.makedirs(temp_dir, exist_ok=True)
    cleanup_old_files(temp_dir)

    file_id = str(uuid.uuid4())
    output_template = os.path.join(temp_dir, f"{file_id}.%(ext)s")

    options = get_ytdl_base_options(url)
    options['outtmpl'] = output_template

    if format_type == 'Video + Audio':
        options['format'] = format_id
        options['merge_output_format'] = 'mp4'
    elif format_type == 'Audio Only':
        options['format'] = format_id
        options['postprocessors'] = [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }]

    if download_id:
        cache.set(f"dl_prog_{download_id}", {'status': 'downloading', 'percent': 0}, timeout=600)

        def progress_hook(d):
            if d['status'] == 'downloading':
                total = d.get('total_bytes') or d.get('total_bytes_estimate') or 0
                downloaded = d.get('downloaded_bytes', 0)
                percent = min(99, round((downloaded / total) * 100)) if total > 0 else 0
                cache.set(f"dl_prog_{download_id}", {
                    'status': 'downloading',
                    'percent': percent,
                    'speed': d.get('speed'),
                    'eta': d.get('eta'),
                }, timeout=600)
            elif d['status'] == 'finished':
                state = 'Merging...' if format_type == 'Video + Audio' else 'Converting to MP3...'
                cache.set(f"dl_prog_{download_id}", {
                    'status': state,
                    'percent': 99,
                }, timeout=600)

        options['progress_hooks'] = [progress_hook]

    def _download(opts):
        with yt_dlp.YoutubeDL(opts) as ydl:
            return ydl.extract_info(url, download=True)

    try:
        info = _execute_with_retry(_download, url, options)

        valid_files = [
            os.path.join(temp_dir, f)
            for f in os.listdir(temp_dir)
            if f.startswith(file_id)
            and not f.endswith('.part')
            and '.f' not in f
            and not f.endswith('.ytdl')
        ]

        if not valid_files:
            raise Exception("Failed to locate final downloaded file")

        valid_files.sort(key=lambda x: len(x))
        downloaded_file = valid_files[0]

        if format_type == 'Video + Audio':
            has_video, has_audio, audio_codec, video_codec = verify_video_audio_streams(downloaded_file)
            if not (has_video and has_audio):
                _safe_remove(downloaded_file)
                raise Exception("FFmpeg merge failed – missing audio/video track.")

            ffmpeg_cmd = _ffmpeg_cmd()
            transcoded_file = os.path.join(temp_dir, f"{file_id}_transcoded.mp4")

            is_h264 = video_codec and ('h264' in video_codec or 'avc' in video_codec)
            video_codec_arg = (
                ['-c:v', 'copy'] if is_h264
                else ['-c:v', 'libx264', '-preset', 'superfast', '-crf', '23']
            )

            cmd = (
                [ffmpeg_cmd, '-i', downloaded_file]
                + video_codec_arg
                + ['-c:a', 'aac', '-b:a', '192k', '-ar', '44100', '-ac', '2',
                   '-movflags', '+faststart', transcoded_file, '-y']
            )

            logger.info("Transcoding video (%s) audio (%s) to H.264/AAC", video_codec, audio_codec)
            r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if r.returncode == 0 and os.path.exists(transcoded_file):
                _safe_remove(downloaded_file)
                downloaded_file = transcoded_file
            else:
                logger.error("Transcode failed: %s", r.stderr[:500])
                _safe_remove(downloaded_file)
                raise Exception("Failed to transcode to H.264/AAC.")

        elif format_type == 'Audio Only':
            has_video, has_audio, audio_codec, video_codec = verify_video_audio_streams(downloaded_file)
            if not has_audio:
                _safe_remove(downloaded_file)
                raise Exception("Downloaded file does not contain an audio track.")

            ffmpeg_cmd = _ffmpeg_cmd()
            transcoded_file = os.path.join(temp_dir, f"{file_id}_transcoded.mp3")

            cmd = [
                ffmpeg_cmd, '-i', downloaded_file,
                '-c:a', 'libmp3lame', '-b:a', '192k', '-ar', '44100', '-ac', '2',
                '-vn', transcoded_file, '-y',
            ]

            logger.info("Transcoding audio (%s) to MP3", audio_codec)
            r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if r.returncode == 0 and os.path.exists(transcoded_file):
                _safe_remove(downloaded_file)
                downloaded_file = transcoded_file
            else:
                logger.error("Audio transcode failed: %s", r.stderr[:500])
                _safe_remove(downloaded_file)
                raise Exception("Failed to transcode audio to MP3.")

        return downloaded_file, info.get('title', 'video')

    except YTDLPError:
        raise
    except Exception as e:
        logger.error("Download error for domain %s: %s", urlparse(url).netloc, str(e)[:200])
        raise YTDLPError('DOWNLOAD_FAILED', _ERROR_MESSAGES['DOWNLOAD_FAILED'])


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _ffmpeg_cmd():
    """Return the ffmpeg binary path."""
    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
        for name in ('ffmpeg', 'ffmpeg.exe'):
            candidate = os.path.join(str(ffmpeg_dir), name)
            if os.path.isfile(candidate):
                return candidate
    return 'ffmpeg'


def _safe_remove(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def cleanup_old_files(directory, max_age_seconds=600):
    """Delete files older than max_age_seconds from directory."""
    try:
        now = time.time()
        for filename in os.listdir(directory):
            filepath = os.path.join(directory, filename)
            if os.path.isfile(filepath):
                if os.stat(filepath).st_mtime < now - max_age_seconds:
                    _safe_remove(filepath)
    except Exception as e:
        logger.error("Error cleaning up old files in %s: %s", directory, e)

