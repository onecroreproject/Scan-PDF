import os
import time
import uuid
import logging
import subprocess
import json
from django.conf import settings
from urllib.parse import urlparse
try:
    import yt_dlp
except ImportError:
    yt_dlp = None

logger = logging.getLogger(__name__)

import random

def get_ytdl_base_options():
    options = {
        'quiet': True,
        'no_warnings': True,
        'nocheckcertificate': True,
        'geo_bypass': True,
        'retries': 10,
        'fragment_retries': 10,
        'format_sort': ['vcodec:h264', 'res', 'acodec:m4a'],
        'force_ipv4': True,
        'extractor_args': {'youtube': ['player_client=web,default']},
    }
    
    # Secure Optional Authentication Support
    cookies_file = os.environ.get('YTDLP_COOKIE_FILE')
    if cookies_file and os.path.exists(cookies_file):
        options['cookiefile'] = cookies_file
        logger.info("Loaded secure cookies file from environment variable.")

    browser_name = os.environ.get('YTDLP_COOKIES_FROM_BROWSER')
    if browser_name and not options.get('cookiefile'):
        options['cookiesfrombrowser'] = (browser_name.strip(), None, None, None)
        logger.info("Configured yt-dlp to load cookies from %s.", browser_name.strip())
    
    # Try to use bundled FFmpeg if available
    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
        options['ffmpeg_location'] = str(ffmpeg_dir)
        
    return options

class YTDLPError(Exception):
    def __init__(self, code, message, raw_error=None):
        self.code = code
        self.message = message
        self.raw_error = raw_error
        super().__init__(self.message)

def _categorize_error(e, url):
    error_msg = str(e).lower()

    if 'facebook.com' in urlparse(url).netloc.lower():
        if any(k in error_msg for k in ['could not copy chrome cookie database', 'cookie database', 'permission denied']):
            return 'FACEBOOK_BROWSER_LOCKED', 'Close Chrome completely, restart this server, and try again. For a running browser, export Facebook cookies and configure YTDLP_COOKIE_FILE instead.'
        if any(k in error_msg for k in ['cannot parse data', 'login', 'sign in', 'private']):
            return 'FACEBOOK_ACCESS_REQUIRED', 'Facebook did not provide the video to this server. Export your Facebook cookies and configure YTDLP_COOKIE_FILE, then try again.'
    
    if "nonetype" in error_msg and "youtubedl" in error_msg:
        return "YT_DLP_MISSING", "yt-dlp is not installed or available on this server."
        
    if "http error 400" in error_msg or "bad request" in error_msg:
        return "YOUTUBE_API_BLOCK", "YouTube API rejected the request. Please update yt-dlp or check player clients."
    
    if any(k in error_msg for k in ['sign in', 'bot', 'age', 'verify', 'cookies-from-browser', 'authentication', 'logged-in', 'login', 'empty media response']):
        return "YOUTUBE_BOT_CHALLENGE", "Unable to analyze this video from the current server. Please try again later."
    
    if any(k in error_msg for k in ['video unavailable', 'unavailable video', 'not available']):
        return "VIDEO_UNAVAILABLE", "This video is unavailable or cannot be accessed from the server."
        
    if any(k in error_msg for k in ['private video']):
        return "AUTH_REQUIRED", "This video requires authorized access."
        
    if any(k in error_msg for k in ['http error 429', 'too many requests']):
        return "RATE_LIMITED", "YouTube is temporarily limiting requests. Please try again shortly."
        
    if any(k in error_msg for k in ['network', 'timeout', 'timed out', 'connection']):
        return "NETWORK_ERROR", "A network error occurred while reaching the video server."
        
    if any(k in error_msg for k in ['format is not available', 'requested format']):
        return "FORMAT_UNAVAILABLE", "The requested format is not available."
    if any(k in error_msg for k in ['ffmpeg is not installed', 'ffmpeg not found', 'ffprobe and ffmpeg']):
        return "FFMPEG_MISSING", "FFmpeg is required for this format but is not available on the server."
        
    return "UNKNOWN_YOUTUBE_ERROR", "Unable to prepare this download."

def _execute_with_retry(execute_func, url, options):
    """
    Executes a yt-dlp function. Removes unsupported headless browser retries.
    """
    try:
        return execute_func(options)
    except Exception as e:
        import traceback
        code, safe_msg = _categorize_error(e, url)
        raw_err_str = str(e) + "\n\n" + traceback.format_exc()
        logger.error(
            "YouTube analysis/download failed.",
            extra={
                "video_url": url,
                "error_type": code,
                "yt_dlp_version": getattr(yt_dlp.version, '__version__', 'unknown') if hasattr(yt_dlp, 'version') else getattr(yt_dlp, '__version__', 'unknown'),
                "raw_error": str(e),
                "traceback": traceback.format_exc()
            }
        )
        raise YTDLPError(code, safe_msg, raw_error=raw_err_str)


def verify_video_audio_streams(filepath):
    """
    Uses ffprobe to verify that the file contains both a video and an audio stream.
    """
    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    ffprobe_cmd = 'ffprobe'
    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
        ffprobe_cmd = os.path.join(ffmpeg_dir, 'ffprobe')
        
    cmd = [
        ffprobe_cmd,
        '-v', 'quiet',
        '-print_format', 'json',
        '-show_streams',
        filepath
    ]
    
    try:
        # Use CREATE_NO_WINDOW on Windows to prevent popups, but subprocess.run is usually fine
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode != 0:
            logger.error(f"ffprobe failed for {filepath}: {result.stderr}")
            return False
            
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
        logger.error(f"Error running ffprobe on {filepath}: {e}")
        return False, False, None

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

def analyze_video(url):
    """
    Analyzes the given URL using yt-dlp and returns available formats and metadata.
    """
    options = get_ytdl_base_options()
    
    def _extract(opts):
        with yt_dlp.YoutubeDL(opts) as ydl:
            return ydl.extract_info(url, download=False)
            
    try:
        info = _execute_with_retry(_extract, url, options)
        
        # Basic metadata
        if True:
            # Extract duration safely
            duration = info.get('duration')
            if not duration:
                for f in info.get('formats', []):
                    if f.get('duration'):
                        duration = f.get('duration')
                        break
            if not duration:
                duration = 0
                
            result = {
                'title': info.get('title', 'Unknown Title'),
                'thumbnail': info.get('thumbnail'),
                'duration': duration,
                'uploader': info.get('uploader', info.get('extractor_key')),
                'formats': []
            }
            
            # Standard parsing for YouTube and all platforms
            video_formats_by_height = {}
            audio_formats = []
            
            for f in info.get('formats', []):
                vcodec = str(f.get('vcodec') or 'none').lower()
                acodec = str(f.get('acodec') or 'none').lower()
                ext = str(f.get('ext') or 'unknown').lower()
                
                filesize = f.get('filesize') or f.get('filesize_approx') or 0
                height = f.get('height') or 0
                width = f.get('width') or 0
                bitrate = f.get('tbr') or f.get('vbr') or f.get('abr') or 0
                
                # Some social extractors (Instagram, TikTok) don't label vcodec but output MP4s
                if vcodec == 'none' and acodec == 'none':
                    if (height > 0 or width > 0) and ext in ['mp4', 'webm', 'mov']:
                        vcodec = 'unknown' # Force parsing as video
                    else:
                        continue
                
                if vcodec == 'none' and acodec != 'none':
                    abr = f.get('abr') or bitrate or 0
                    approx_bitrate = round(abr / 32) * 32 if abr else 0
                    audio_formats.append({
                        'format_id': f.get('format_id'),
                        'resolution': f.get('format_note', 'Audio'),
                        'ext': 'mp3',
                        'vcodec': 'none',
                        'acodec': 'MP3',
                        'bitrate': bitrate,
                        'filesize': filesize,
                        'abr': approx_bitrate,
                        'type': 'Audio Only'
                    })
                    continue
                
                if height > 0:
                    priority = get_codec_priority(vcodec)
                    
                    # Normalize vertical video resolution (Shorts)
                    display_height = width if (height > width and width > 0) else height
                    
                    fmt = {
                        'format_id': f.get('format_id'),
                        'resolution': f"{display_height}p",
                        'ext': 'mp4', # Force MP4
                        'vcodec': 'H.264' if priority >= 3 else vcodec, # Simplify UI
                        'acodec': 'AAC',
                        'bitrate': bitrate,
                        'filesize': filesize,
                        'priority': priority,
                        'type': 'Video + Audio',
                        'raw_acodec': acodec
                    }
                    
                    if display_height not in video_formats_by_height:
                        video_formats_by_height[display_height] = fmt
                    else:
                        if priority > video_formats_by_height[display_height]['priority']:
                            video_formats_by_height[display_height] = fmt
            
            final_video_formats = []
            for height in sorted(video_formats_by_height.keys(), reverse=True):
                fmt = video_formats_by_height[height]
                final_video_formats.append(fmt)
                
            seen_abr = set()
            final_audio_formats = []
            
            audio_formats.sort(key=lambda x: x['abr'], reverse=True)
            for fmt in audio_formats:
                if fmt['abr'] not in seen_abr:
                    seen_abr.add(fmt['abr'])
                    final_audio_formats.append(fmt)
                    
            if not final_audio_formats:
                final_audio_formats.append({
                    'format_id': 'bestaudio/best',
                    'resolution': 'Best Audio',
                    'ext': 'mp3',
                    'vcodec': 'none',
                    'acodec': 'MP3',
                    'bitrate': 0,
                    'filesize': 0,
                    'type': 'Audio Only'
                })
                
            if not final_video_formats:
                final_video_formats.append({
                    'format_id': 'bestvideo+bestaudio/best',
                    'resolution': 'Highest Quality',
                    'ext': 'mp4',
                    'vcodec': 'H.264',
                    'acodec': 'AAC',
                    'bitrate': 0,
                    'filesize': 0,
                    'type': 'Video + Audio'
                })
                
            result['formats'] = final_video_formats + final_audio_formats
            return result
    except YTDLPError:
        raise
    except Exception as e:
        import traceback
        logger.error(f"Failed to analyze URL {url}: {e}")
        raw_err_str = str(e) + "\n\n" + traceback.format_exc()
        raise YTDLPError("METADATA_PARSE_FAILED", "Failed to parse video metadata.", raw_error=raw_err_str)

def download_format(url, format_id, format_type):
    """
    Downloads the specific format.
    Returns the absolute path to the downloaded file.
    """
    import shutil
    
    # Format IDs can change between metadata extraction and download, especially
    # for Facebook. Let yt-dlp validate the selected ID during the real download.
    selected_fmt = None
    
    ffmpeg_available = bool(shutil.which('ffmpeg'))
    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    if not ffmpeg_available and ffmpeg_dir and os.path.exists(ffmpeg_dir):
        ffmpeg_available = bool(shutil.which('ffmpeg', path=ffmpeg_dir))

    actual_ytdl_format = format_id
    if format_type == 'Video + Audio':
        if selected_fmt and selected_fmt.get('raw_acodec') == 'none':
            if not ffmpeg_available:
                raise YTDLPError("FFMPEG_REQUIRED", "This quality requires media merging, which is currently unavailable.")
            actual_ytdl_format = f"{format_id}+bestaudio/best"
            
    if format_type == 'Audio Only':
        if not ffmpeg_available:
            raise YTDLPError("FFMPEG_REQUIRED", "Audio extraction requires media merging, which is currently unavailable.")

    temp_dir = os.path.join(settings.MEDIA_ROOT, 'video_downloads')
    os.makedirs(temp_dir, exist_ok=True)
    cleanup_old_files(temp_dir)
    
    file_id = str(uuid.uuid4())
    output_template = os.path.join(temp_dir, f"{file_id}.%(ext)s")
    
    options = get_ytdl_base_options()
    options['outtmpl'] = output_template
    
    if format_type == 'Video + Audio':
        options['format'] = actual_ytdl_format
        options['merge_output_format'] = 'mp4'
    if format_type == 'Audio Only':
        options['format'] = actual_ytdl_format
        options['postprocessors'] = [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }]
        
    def _download(opts):
        with yt_dlp.YoutubeDL(opts) as ydl:
            return ydl.extract_info(url, download=True)
            
    try:
        info = _execute_with_retry(_download, url, options)
        
        # Find the actual downloaded file, ignoring yt-dlp intermediate files
        if True:
            valid_files = []
            for f in os.listdir(temp_dir):
                if f.startswith(file_id) and not f.endswith('.part') and '.f' not in f and not f.endswith('.ytdl'):
                    valid_files.append(os.path.join(temp_dir, f))
                    
            if not valid_files:
                raise Exception("Failed to locate final downloaded file")
                
            # If multiple files exist (rare if cleaned correctly), pick the one without extra extensions,
            # or simply sort by length as intermediate files tend to have longer names.
            valid_files.sort(key=lambda x: len(x))
            downloaded_file = valid_files[0]
            
            # Verify Video + Audio merge
            if format_type == 'Video + Audio':
                has_video, has_audio, audio_codec, video_codec = verify_video_audio_streams(downloaded_file)
                if not (has_video and has_audio):
                    # Delete the defective file
                    try:
                        os.remove(downloaded_file)
                    except Exception:
                        pass
                    raise Exception("FFmpeg merge failed or resulted in a silent video. Missing audio track.")
                    
                # ALWAYS transcode to ensure compatibility with Windows Media Player / browsers
                if True:
                    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
                    ffmpeg_cmd = 'ffmpeg'
                    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
                        ffmpeg_cmd = os.path.join(ffmpeg_dir, 'ffmpeg')
                        
                    transcoded_file = os.path.join(temp_dir, f"{file_id}_transcoded.mp4")
                    
                    # If video is already H.264/AVC, copy it (fast); otherwise transcode it to libx264
                    is_h264 = video_codec and ('h264' in video_codec or 'avc' in video_codec)
                    video_codec_arg = ['-c:v', 'copy'] if is_h264 else ['-c:v', 'libx264', '-preset', 'superfast', '-crf', '23']
                    
                    cmd = [
                        ffmpeg_cmd,
                        '-i', downloaded_file,
                    ] + video_codec_arg + [
                        '-c:a', 'aac',
                        '-b:a', '192k',
                        '-ar', '44100',
                        '-ac', '2',
                        '-movflags', '+faststart',
                        transcoded_file,
                        '-y'
                    ]
                    
                    logger.info(f"Transcoding video ({video_codec}) and audio ({audio_codec}) to standard H.264/AAC: {' '.join(cmd)}")
                    # Use CREATE_NO_WINDOW if available to prevent popup
                    transcode_result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                    if transcode_result.returncode == 0 and os.path.exists(transcoded_file):
                        try:
                            os.remove(downloaded_file)
                        except Exception:
                            pass
                        downloaded_file = transcoded_file
                    else:
                        logger.error(f"Transcode failed: {transcode_result.stderr}")
                        # Even if transcode fails, we might just return the original or throw error.
                        # Since user wants STRICT compatibility, we should probably raise an error
                        try:
                            os.remove(downloaded_file)
                        except Exception:
                            pass
                        raise Exception("Failed to convert unsupported audio codec to AAC.")
            elif format_type == 'Audio Only':
                has_video, has_audio, audio_codec, video_codec = verify_video_audio_streams(downloaded_file)
                if not has_audio:
                    # Delete the defective file
                    try:
                        os.remove(downloaded_file)
                    except Exception:
                        pass
                    raise Exception("Downloaded file does not contain an audio track.")
                    
                ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
                ffmpeg_cmd = 'ffmpeg'
                if ffmpeg_dir and os.path.exists(ffmpeg_dir):
                    ffmpeg_cmd = os.path.join(ffmpeg_dir, 'ffmpeg')
                    
                transcoded_file = os.path.join(temp_dir, f"{file_id}_transcoded.mp3")
                
                cmd = [
                    ffmpeg_cmd,
                    '-i', downloaded_file,
                    '-c:a', 'libmp3lame',
                    '-b:a', '192k',
                    '-ar', '44100',
                    '-ac', '2',
                    '-vn',
                    transcoded_file,
                    '-y'
                ]
                
                logger.info(f"Transcoding audio ({audio_codec}) to standard MP3: {' '.join(cmd)}")
                transcode_result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                if transcode_result.returncode == 0 and os.path.exists(transcoded_file):
                    try:
                        os.remove(downloaded_file)
                    except Exception:
                        pass
                    downloaded_file = transcoded_file
                else:
                    logger.error(f"Audio Transcode failed: {transcode_result.stderr}")
                    try:
                        os.remove(downloaded_file)
                    except Exception:
                        pass
                    raise Exception("Failed to convert audio stream to MP3 format.")
                
            return downloaded_file, info.get('title', 'video')
            
    except YTDLPError:
        raise
    except Exception as e:
        logger.error(f"Error downloading video URL {url} format {format_id}: {e}")
        raise YTDLPError("DOWNLOAD_FAILED", f"Failed to download file: {str(e)}")

def cleanup_old_files(directory, max_age_seconds=600):
    """Deletes files older than max_age_seconds in the given directory."""
    try:
        now = time.time()
        for filename in os.listdir(directory):
            filepath = os.path.join(directory, filename)
            if os.path.isfile(filepath):
                if os.stat(filepath).st_mtime < now - max_age_seconds:
                    try:
                        os.remove(filepath)
                    except Exception as e:
                        logger.warning(f"Could not remove old file {filepath}: {e}")
    except Exception as e:
        logger.error(f"Error cleaning up old files in {directory}: {e}")
