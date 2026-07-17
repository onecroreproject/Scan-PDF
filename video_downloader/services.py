import os
import time
import uuid
import logging
import subprocess
import json
from django.conf import settings
from urllib.parse import urlparse
import yt_dlp

logger = logging.getLogger(__name__)

def get_ytdl_base_options():
    options = {
        'quiet': True,
        'no_warnings': True,
        'nocheckcertificate': True,
        'geo_bypass': True,
        'retries': 10,
        'fragment_retries': 10,
        'http_headers': {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
    }
    
    # Try to use bundled FFmpeg if available
    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
        options['ffmpeg_location'] = str(ffmpeg_dir)
        
    return options

def _is_bot_error(e, url):
    error_msg = str(e).lower()
    bot_keywords = [
        'sign in', 'bot', 'age', 'verify', 
        'cookies-from-browser', 'authentication', 
        'logged-in', 'empty media response', 'login'
    ]
    return any(keyword in error_msg for keyword in bot_keywords)

def _execute_with_retry(execute_func, url, options):
    """
    Executes a yt-dlp function. If it fails due to YouTube bot detection,
    retries with browser cookies in order: Chrome, Edge, Firefox.
    """
    try:
        return execute_func(options)
    except Exception as e:
        if not _is_bot_error(e, url):
            raise
            
        logger.warning(f"Bot/age detection encountered for {url}. Retrying with browser cookies...")
        browsers = ['chrome', 'edge', 'firefox']
        
        for browser in browsers:
            retry_options = options.copy()
            retry_options['cookiesfrombrowser'] = [(browser, None, None, None)]
            
            try:
                logger.info(f"Retrying with {browser} cookies...")
                return execute_func(retry_options)
            except Exception as retry_e:
                logger.warning(f"{browser} cookie retry failed: {retry_e}")
                
        raise ValueError("The platform is blocking access. Please ensure you are logged in on Chrome/Edge/Firefox to bypass bot protection, or try again later.")


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
        
        for s in streams:
            codec_type = s.get('codec_type')
            if codec_type == 'video':
                has_video = True
            elif codec_type == 'audio':
                has_audio = True
                audio_codec = s.get('codec_name', '').lower()
                
        return has_video, has_audio, audio_codec
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
            result = {
                'title': info.get('title', 'Unknown Title'),
                'thumbnail': info.get('thumbnail'),
                'duration': info.get('duration'),
                'uploader': info.get('uploader', info.get('extractor_key')),
                'formats': []
            }
            
            extractor = str(info.get('extractor_key', '')).lower()
            social_extractors = ['instagram', 'facebook', 'twitter', 'tiktok', 'x', 'vimeo']
            
            if any(ext in extractor for ext in social_extractors):
                # For social media, provide only best synthetic formats
                result['formats'] = [
                    {
                        'format_id': 'bestvideo+bestaudio/best',
                        'resolution': 'Highest Quality',
                        'ext': 'mp4',
                        'vcodec': 'H.264',
                        'acodec': 'AAC',
                        'filesize': 0,
                        'type': 'Video + Audio'
                    },
                    {
                        'format_id': 'bestaudio/best',
                        'resolution': 'Best Audio',
                        'ext': 'mp3',
                        'vcodec': 'none',
                        'acodec': 'MP3',
                        'filesize': 0,
                        'type': 'Audio Only'
                    }
                ]
                return result

            # Standard parsing for YouTube and others
            video_formats_by_height = {}
            audio_formats = []
            
            for f in info.get('formats', []):
                vcodec = f.get('vcodec', 'none')
                acodec = f.get('acodec', 'none')
                
                if vcodec == 'none' and acodec == 'none':
                    continue
                    
                filesize = f.get('filesize') or f.get('filesize_approx') or 0
                height = f.get('height') or 0
                bitrate = f.get('tbr') or f.get('vbr') or f.get('abr') or 0
                ext = f.get('ext', 'unknown')
                
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
                    
                    fmt = {
                        'format_id': f.get('format_id'),
                        'resolution': f"{height}p",
                        'ext': 'mp4', # Force MP4
                        'vcodec': 'H.264' if priority >= 3 else vcodec, # Simplify UI
                        'acodec': 'AAC',
                        'bitrate': bitrate,
                        'filesize': filesize,
                        'priority': priority,
                        'type': 'Video + Audio',
                        'raw_acodec': acodec
                    }
                    
                    if height not in video_formats_by_height:
                        video_formats_by_height[height] = fmt
                    else:
                        if priority > video_formats_by_height[height]['priority']:
                            video_formats_by_height[height] = fmt
            
            final_video_formats = []
            for height in sorted(video_formats_by_height.keys(), reverse=True):
                fmt = video_formats_by_height[height]
                original_id = fmt['format_id']
                if fmt['raw_acodec'] == 'none':
                    fmt['format_id'] = f"{original_id}+bestaudio/best"
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
    except Exception as e:
        logger.error(f"Error analyzing video URL {url}: {e}")
        raise ValueError(str(e))

def download_format(url, format_id, format_type):
    """
    Downloads the specific format.
    Returns the absolute path to the downloaded file.
    """
    temp_dir = os.path.join(settings.MEDIA_ROOT, 'video_downloads')
    os.makedirs(temp_dir, exist_ok=True)
    cleanup_old_files(temp_dir)
    
    file_id = str(uuid.uuid4())
    output_template = os.path.join(temp_dir, f"{file_id}.%(ext)s")
    
    options = get_ytdl_base_options()
    options['outtmpl'] = output_template
    
    if format_type == 'Video + Audio':
        options['format'] = format_id
        options['merge_output_format'] = 'mp4'
    if format_type == 'Audio Only':
        options['format'] = format_id
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
                has_video, has_audio, audio_codec = verify_video_audio_streams(downloaded_file)
                if not (has_video and has_audio):
                    # Delete the defective file
                    try:
                        os.remove(downloaded_file)
                    except Exception:
                        pass
                    raise Exception("FFmpeg merge failed or resulted in a silent video. Missing audio track.")
                    
                # Transcode unsupported audio to AAC if necessary
                unsupported_codecs = ['opus', 'vorbis', 'webm']
                if audio_codec in unsupported_codecs:
                    ffmpeg_dir = getattr(settings, 'FFMPEG_BIN_DIR', None)
                    ffmpeg_cmd = 'ffmpeg'
                    if ffmpeg_dir and os.path.exists(ffmpeg_dir):
                        ffmpeg_cmd = os.path.join(ffmpeg_dir, 'ffmpeg')
                        
                    transcoded_file = os.path.join(temp_dir, f"{file_id}_aac.mp4")
                    cmd = [
                        ffmpeg_cmd,
                        '-i', downloaded_file,
                        '-c:v', 'copy',
                        '-c:a', 'aac',
                        '-b:a', '192k',
                        transcoded_file,
                        '-y'
                    ]
                    
                    logger.info(f"Transcoding audio from {audio_codec} to AAC: {' '.join(cmd)}")
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
                
            return downloaded_file, info.get('title', 'video')
            
    except Exception as e:
        logger.error(f"Error downloading video URL {url} format {format_id}: {e}")
        raise ValueError(str(e))

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
