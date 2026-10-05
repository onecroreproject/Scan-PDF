import yt_dlp
import logging
from urllib.parse import urlparse
import tempfile
import os
import uuid
import re

logger = logging.getLogger(__name__)

def sanitize_filename(name):
    clean = re.sub(r'[^a-zA-Z0-9\.\_\-]', '_', name)
    return clean[:200]

SUPPORTED_PLATFORMS = {
    'instagram': ['instagram.com', 'www.instagram.com'],
    'twitter': ['twitter.com', 'www.twitter.com', 'x.com', 'www.x.com'],
    'facebook': ['facebook.com', 'www.facebook.com', 'fb.watch'],
    'threads': ['threads.net', 'www.threads.net', 'threads.com', 'www.threads.com'],
    'youtube': ['youtube.com', 'www.youtube.com', 'youtu.be']
}

def is_supported_url(url):
    try:
        parsed = urlparse(url)
        if parsed.scheme not in ('http', 'https'):
            return None
        
        domain = parsed.hostname
        if not domain:
            return None
            
        domain = domain.lower()
        if domain.startswith('www.'):
            domain = domain[4:]
        
        for platform, domains in SUPPORTED_PLATFORMS.items():
            if domain in domains or any(domain.endswith('.' + d) for d in domains):
                return platform
        return None
    except Exception:
        return None

def get_media_info(url):
    platform = is_supported_url(url)
    if not platform:
        raise ValueError("Unsupported platform or invalid URL.")

    ydl_opts = {
        'quiet': True,
        'no_warnings': True,
        'skip_download': True,
        'geo_bypass': True,
        'extract_flat': 'in_playlist',
    }

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
            
            # Normalize the result
            items = []
            title = info.get('title', 'Media Download')
            thumbnail = info.get('thumbnail')
            author = info.get('uploader') or info.get('extractor_key')
            
            # Handle formats
            if 'formats' in info:
                for f in info['formats']:
                    # Only include formats that have video or image and a valid URL
                    if f.get('url') and (f.get('vcodec') != 'none' or f.get('acodec') != 'none'):
                        format_type = "video" if f.get('vcodec') != 'none' else "audio"
                        
                        # Sometimes we just have images
                        if f.get('ext') in ['jpg', 'png', 'webp']:
                            format_type = "image"
                        
                        quality = f.get('format_note') or f.get('resolution') or f"{f.get('height', 0)}p" if f.get('height') else "Best"
                        has_audio = f.get('acodec') != 'none'
                        items.append({
                            'id': f.get('format_id'),
                            'type': format_type,
                            'quality': quality,
                            'format': f.get('ext', 'unknown'),
                            'width': f.get('width'),
                            'height': f.get('height'),
                            'fileSize': f.get('filesize') or f.get('filesize_approx'),
                            'has_audio': has_audio,
                            'can_download_mp3': has_audio or format_type == 'audio',
                            'url': f.get('url')  # we return the direct URL for streaming
                        })
                        
            # Some extractors don't provide formats list but direct URL
            if not items and info.get('url'):
                ext = info.get('ext', 'mp4')
                format_type = "image" if ext in ['jpg', 'png', 'webp'] else "video"
                items.append({
                    'id': 'default',
                    'type': format_type,
                    'quality': 'Best available',
                    'format': ext,
                    'fileSize': info.get('filesize') or info.get('filesize_approx'),
                    'has_audio': format_type != 'image',
                    'can_download_mp3': format_type != 'image',
                    'url': info.get('url')
                })

            # If playlist / gallery (like multiple instagram slides)
            if 'entries' in info:
                for entry in info['entries']:
                    if entry.get('url'):
                        ext = entry.get('ext', 'mp4')
                        format_type = "image" if ext in ['jpg', 'png', 'webp'] else "video"
                        items.append({
                            'id': entry.get('id', 'default'),
                            'type': format_type,
                            'quality': 'Best available',
                            'format': ext,
                            'fileSize': entry.get('filesize'),
                            'has_audio': format_type != 'image',
                            'can_download_mp3': format_type != 'image',
                            'url': entry.get('url')
                        })

            # Filter duplicates and bad formats
            unique_video_items = {}
            has_any_audio = False
            image_items = []

            for item in items:
                if item.get('has_audio'):
                    has_any_audio = True

                if item['type'] == 'image':
                    # Avoid duplicates
                    if not any(img['url'] == item['url'] for img in image_items):
                        image_items.append(item)
                elif item['type'] == 'video':
                    if 'm3u8' in item.get('format', ''):
                        continue
                    q = item['quality']
                    if q not in unique_video_items:
                        unique_video_items[q] = item
                    else:
                        # Prefer mp4
                        curr = unique_video_items[q]
                        if item['format'] == 'mp4' and curr['format'] != 'mp4':
                            unique_video_items[q] = item
                        elif item['format'] == curr['format']:
                            if (item.get('fileSize') or 0) > (curr.get('fileSize') or 0):
                                unique_video_items[q] = item
                elif item['type'] == 'audio':
                    has_any_audio = True

            video_formats = list(unique_video_items.values())

            # Sort items by height (quality) descending
            def get_height(x):
                h = x.get('height')
                return h if h is not None else 0
            
            video_formats.sort(key=get_height, reverse=True)

            audio_info = {
                'available': has_any_audio,
                'qualities': [
                    {'label': 'Best / 320 kbps', 'value': '320'},
                    {'label': '256 kbps', 'value': '256'},
                    {'label': '192 kbps', 'value': '192'},
                    {'label': '128 kbps', 'value': '128'},
                ] if has_any_audio else []
            }

            return {
                'platform': platform,
                'title': title,
                'author': author,
                'thumbnail': thumbnail,
                'duration': info.get('duration'),
                'video_formats': video_formats,
                'audio': audio_info,
                'image_formats': image_items
            }
    except Exception as e:
        import re
        raw_msg = str(e)
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        clean_msg = ansi_escape.sub('', raw_msg)
        
        logger.error(f"yt-dlp info error: {clean_msg}")
        
        if "Unsupported URL" in clean_msg or "not supported" in clean_msg.lower():
            raise ValueError("Unable to access this public Threads post.")
        elif "Sign in" in clean_msg or "Private" in clean_msg or "login" in clean_msg.lower() or "401" in clean_msg or "403" in clean_msg:
            raise ValueError("This Threads post is not publicly accessible.")
        else:
            raise ValueError("No downloadable media was found in this Threads post.")

import os
import shutil

def get_ffmpeg_location():
    from django.conf import settings
    configured = getattr(settings, "FFMPEG_LOCATION", None)
    
    # Check old settings variables too just in case
    if not configured:
        configured = getattr(settings, "FFMPEG_BIN_DIR", None) or getattr(settings, "FFMPEG_PATH", None)

    if configured:
        configured = os.path.abspath(configured)

        if os.path.isfile(configured):
            configured = os.path.dirname(configured)

        ffmpeg_exe = os.path.join(configured, "ffmpeg.exe")
        ffprobe_exe = os.path.join(configured, "ffprobe.exe")

        if os.path.isfile(ffmpeg_exe) and os.path.isfile(ffprobe_exe):
            return configured
        elif os.path.isfile(os.path.join(configured, "ffmpeg")) and os.path.isfile(os.path.join(configured, "ffprobe")):
            return configured

    ffmpeg_path = shutil.which("ffmpeg")
    ffprobe_path = shutil.which("ffprobe")

    if not ffmpeg_path or not ffprobe_path:
        return None

    ffmpeg_dir = os.path.dirname(os.path.abspath(ffmpeg_path))
    ffprobe_dir = os.path.dirname(os.path.abspath(ffprobe_path))

    if os.path.normcase(ffmpeg_dir) == os.path.normcase(ffprobe_dir):
        return ffmpeg_dir

    return ffmpeg_dir

def is_ffmpeg_available():
    loc = get_ffmpeg_location()
    logger.info(f"FFmpeg detected: {loc is not None}")
    if loc:
        logger.info(f"FFmpeg directory: {loc}")
        logger.info(f"FFprobe detected: True")
    return loc is not None

def download_media_to_temp(url, format_id, download_type="video", audio_quality="192"):
    platform = is_supported_url(url)
    if not platform:
        raise ValueError("Unsupported platform or invalid URL.")

    temp_dir = tempfile.gettempdir()
    unique_filename = f"dl_{uuid.uuid4().hex}"
    out_tmpl = os.path.join(temp_dir, f"{unique_filename}.%(ext)s")

    ydl_opts = {
        'quiet': True,
        'no_warnings': True,
        'outtmpl': out_tmpl,
    }
    
    ffmpeg_loc = get_ffmpeg_location()

    if download_type == 'audio':
        if not is_ffmpeg_available():
            raise ValueError("MP3 conversion requires FFmpeg on the server.")
            
        ydl_opts['format'] = 'bestaudio/best'
        ydl_opts['postprocessors'] = [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': str(audio_quality),
        }]
        
        if ffmpeg_loc:
            ydl_opts['ffmpeg_location'] = ffmpeg_loc
    else:
        target_format = f"{format_id}+bestaudio/bestaudio+{format_id}/{format_id}/best" if format_id and format_id != 'default' else 'bestvideo+bestaudio/best'
        
        if not is_ffmpeg_available():
            if '+' in target_format or 'bestvideo' in target_format:
                raise ValueError("This video quality requires FFmpeg to merge video and audio.")
            ydl_opts['format'] = format_id if format_id and format_id != 'default' else 'best'
        else:
            ydl_opts['format'] = target_format
            if ffmpeg_loc:
                ydl_opts['ffmpeg_location'] = ffmpeg_loc

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            
            if download_type == 'audio':
                filepath = os.path.join(temp_dir, f"{unique_filename}.mp3")
                ext = 'mp3'
            else:
                ext = info.get('ext', 'mp4')
                filepath = os.path.join(temp_dir, f"{unique_filename}.{ext}")
            
            if not os.path.exists(filepath):
                for f in os.listdir(temp_dir):
                    if f.startswith(unique_filename):
                        filepath = os.path.join(temp_dir, f)
                        ext = f.split('.')[-1]
                        break
            
            if not os.path.exists(filepath):
                raise ValueError("Downloaded file not found. If extracting MP3, ensure FFmpeg is installed.")
                
            title = info.get('title', 'media')
            safe_title = sanitize_filename(title) or 'downloaded_media'
            return filepath, f"{safe_title}.{ext}"
    except yt_dlp.utils.PostProcessingError as e:
        logger.error(f"yt-dlp postprocessing error: {str(e)}")
        raise ValueError("MP3 conversion requires FFmpeg on the server.")
    except Exception as e:
        raw_msg = str(e)
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        clean_msg = ansi_escape.sub('', raw_msg)
        
        logger.error(f"yt-dlp download error: {clean_msg}")
        
        if "403" in clean_msg or "Forbidden" in clean_msg:
            raise ValueError("The platform rejected this download request (HTTP 403 Forbidden). Please try another quality or try again later.")
        elif "ffprobe and ffmpeg not found" in clean_msg.lower():
            raise ValueError("This video quality requires FFmpeg to merge video and audio.")
        elif "No video formats" in clean_msg:
            raise ValueError("No video formats available for the requested quality.")
        else:
            raise ValueError("This media could not be downloaded. Please try another quality.")
