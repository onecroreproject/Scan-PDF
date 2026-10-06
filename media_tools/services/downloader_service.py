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

import urllib.request
import urllib.parse
import json
import html

class SafeThreadsRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parsed = urllib.parse.urlparse(newurl)
        allowed = ['threads.net', 'www.threads.net', 'threads.com', 'www.threads.com']
        if parsed.hostname not in allowed:
            raise ValueError(f"Invalid redirect to {parsed.hostname}")
        if parsed.scheme not in ['http', 'https']:
            raise ValueError("Invalid redirect scheme")
        return super().redirect_request(req, fp, code, msg, headers, newurl)

def get_best_image_candidate(candidates_list):
    if not isinstance(candidates_list, list):
        return None
    valid_cands = [c for c in candidates_list if isinstance(c, dict) and 'url' in c]
    if not valid_cands:
        return None
    sorted_cands = sorted(
        valid_cands,
        key=lambda x: (x.get('width', 0) * x.get('height', 0)),
        reverse=True
    )
    return sorted_cands[0]['url']

def extract_images_from_post_object(post_obj):
    images = []
    if 'carousel_media' in post_obj and isinstance(post_obj['carousel_media'], list):
        for item in post_obj['carousel_media']:
            if isinstance(item, dict) and 'image_versions2' in item and 'candidates' in item['image_versions2']:
                best = get_best_image_candidate(item['image_versions2']['candidates'])
                if best:
                    images.append(best)
    elif 'image_versions2' in post_obj and 'candidates' in post_obj['image_versions2']:
        best = get_best_image_candidate(post_obj['image_versions2']['candidates'])
        if best:
            images.append(best)
    elif 'candidates' in post_obj and isinstance(post_obj['candidates'], list):
        best = get_best_image_candidate(post_obj['candidates'])
        if best:
            images.append(best)
    elif 'display_url' in post_obj and isinstance(post_obj['display_url'], str):
        images.append(post_obj['display_url'])
    elif 'image_url' in post_obj and isinstance(post_obj['image_url'], str):
        images.append(post_obj['image_url'])
    return images

def find_post_object(data, shortcode):
    if isinstance(data, dict):
        if data.get('code') == shortcode or data.get('shortcode') == shortcode:
            return data
        for k, v in data.items():
            if isinstance(v, (dict, list)):
                res = find_post_object(v, shortcode)
                if res:
                    return res
    elif isinstance(data, list):
        for item in data:
            if isinstance(item, (dict, list)):
                res = find_post_object(item, shortcode)
                if res:
                    return res
    return None

def find_any_media_object(data):
    if isinstance(data, dict):
        if 'image_versions2' in data or 'carousel_media' in data or 'candidates' in data:
            return data
        for k, v in data.items():
            if isinstance(v, (dict, list)):
                res = find_any_media_object(v)
                if res:
                    return res
    elif isinstance(data, list):
        for item in data:
            if isinstance(item, (dict, list)):
                res = find_any_media_object(item)
                if res:
                    return res
    return None

def extract_all_json_from_string(s):
    results = []
    start_indices = [m.start() for m in re.finditer(r'\{"', s)]
    for i in start_indices:
        stack = 0
        in_str = False
        escape = False
        for j in range(i, len(s)):
            c = s[j]
            if escape:
                escape = False
                continue
            if c == '\\':
                escape = True
                continue
            if c == '"':
                in_str = not in_str
                continue
            
            if not in_str:
                if c == '{':
                    stack += 1
                elif c == '}':
                    stack -= 1
                    if stack == 0:
                        block = s[i:j+1]
                        if ('"code"' in block or '"shortcode"' in block or '"image_versions2"' in block or '"candidates"' in block):
                            try:
                                obj = json.loads(block)
                                results.append(obj)
                            except:
                                pass
                        break
    return results

def extract_threads_images(original_url):
    import re
    shortcode_match = re.search(r'/(?:t|share|post|@[\w.-]+/post)/([a-zA-Z0-9_-]+)', original_url)
    if shortcode_match:
        original_url = f"https://www.threads.net/t/{shortcode_match.group(1)}"
        
    try:
        opener = urllib.request.build_opener(SafeThreadsRedirectHandler())
        req = urllib.request.Request(
            original_url, 
            headers={
                'User-Agent': 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
            }
        )
        with opener.open(req, timeout=10) as response:
            html_content = response.read().decode('utf-8', errors='ignore')
            final_url = response.geturl()
            status = response.status
            
        candidate_images = []
        parsed_objects = []
        
        # 1. Parse all script blocks for JSON robustly
        script_matches = re.findall(r'<script[^>]*>([\s\S]*?)</script>', html_content)
        for s in script_matches:
            if 'image_versions2' in s or 'candidates' in s or 'display_url' in s or 'carousel_media' in s or 'RelayPreloadedState' in s or 'requireLazy' in s:
                parsed_objects.extend(extract_all_json_from_string(s))

        # 2. Identify the requested post object
        shortcode_match = re.search(r'/(?:t|share|post)/([a-zA-Z0-9_-]+)', original_url)
        shortcode = shortcode_match.group(1) if shortcode_match else None
        
        post_obj = None
        for obj in parsed_objects:
            if shortcode:
                post_obj = find_post_object(obj, shortcode)
                if post_obj:
                    break
                    
        # 3. Extract candidate images
        if post_obj:
            candidate_images.extend(extract_images_from_post_object(post_obj))
        else:
            for obj in parsed_objects:
                any_obj = find_any_media_object(obj)
                if any_obj:
                    candidate_images.extend(extract_images_from_post_object(any_obj))
                    if candidate_images:
                        break
                        
        # 4. Fallback to generic JSON-LD and Meta tags if needed
        if not candidate_images:
            for obj in parsed_objects:
                if isinstance(obj, dict):
                    if obj.get('@type') == 'ImageObject':
                        if 'contentUrl' in obj: candidate_images.append(obj['contentUrl'])
                        elif 'url' in obj: candidate_images.append(obj['url'])
                    elif obj.get('@type') in ['SocialMediaPosting', 'DiscussionForumPosting']:
                        if 'image' in obj:
                            images = obj['image']
                            if isinstance(images, str): candidate_images.append(images)
                            elif isinstance(images, list):
                                for img in images:
                                    if isinstance(img, str): candidate_images.append(img)
                                    elif isinstance(img, dict) and 'url' in img: candidate_images.append(img['url'])

            for match in re.finditer(r'<meta[^>]+>', html_content):
                tag = match.group(0)
                if 'property="og:image"' in tag or 'name="twitter:image"' in tag:
                    content_match = re.search(r'content="([^"]+)"', tag)
                    if content_match:
                        val = html.unescape(content_match.group(1))
                        candidate_images.append(val)
                        
        # 5. Deduplicate and validate
        unique_images = []
        for img in candidate_images:
            if not isinstance(img, str):
                continue
            img = img.replace('\\/', '/')
            if 'profile' in img.lower() or 'logo' in img.lower():
                continue
            if img not in unique_images and img.startswith('http'):
                try:
                    req_img = urllib.request.Request(img, headers={
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                        'Range': 'bytes=0-1023'
                    })
                    with urllib.request.urlopen(req_img, timeout=5) as resp_img:
                        ct = resp_img.headers.get('Content-Type', '')
                        if ct.startswith('image/'):
                            unique_images.append(img)
                except urllib.error.HTTPError as e:
                    # Some CDNs reject Range requests, try without if it fails
                    if e.code in [405, 416, 400]:
                        try:
                            req_img = urllib.request.Request(img, headers={'User-Agent': 'Mozilla/5.0'})
                            with urllib.request.urlopen(req_img, timeout=5) as resp_img:
                                ct = resp_img.headers.get('Content-Type', '')
                                if ct.startswith('image/'):
                                    unique_images.append(img)
                        except:
                            pass
                except:
                    pass
                
        # Diagnostic logging
        script_types = list(set(re.findall(r'<script[^>]*type="([^"]+)"', html_content)))
        media_key_hits = []
        if 'image_versions2' in html_content: media_key_hits.append('image_versions2')
        if 'candidates' in html_content: media_key_hits.append('candidates')
        if 'carousel_media' in html_content: media_key_hits.append('carousel_media')
        
        logger.info(f"Threads image hydration scan: scripts={len(script_matches)} media_key_hits={media_key_hits} requested_post_obj_found={post_obj is not None} candidates={len(candidate_images)}")
                
        if not unique_images:
            logger.error(f"Threads image fallback: original_host={urllib.parse.urlparse(original_url).hostname} final_host={urllib.parse.urlparse(final_url).hostname} status={status} html_length={len(html_content)} script_count={len(script_matches)} requested_obj_found={post_obj is not None} media_keys={media_key_hits} candidate_images={len(candidate_images)} validated_images=0")
            raise ValueError("No downloadable media was found in this Threads post.")
        
        title_match = re.search(r'<meta[^>]+property="og:title"[^>]+content="([^"]+)"', html_content)
        if not title_match:
            title_match = re.search(r'<meta[^>]+content="([^"]+)"[^>]+property="og:title"', html_content)
        title = html.unescape(title_match.group(1)) if title_match else "Threads Image"
        
        image_items = []
        for idx, img_url in enumerate(unique_images):
            image_items.append({
                'id': f'img_{idx}',
                'type': 'image',
                'quality': 'Best available',
                'format': 'jpg',
                'fileSize': None,
                'has_audio': False,
                'can_download_mp3': False,
                'url': img_url
            })
            
        return {
            'platform': 'threads',
            'title': title,
            'author': 'Threads',
            'thumbnail': unique_images[0],
            'duration': None,
            'video_formats': [],
            'audio': {'available': False, 'qualities': []},
            'image_formats': image_items
        }
    except Exception as e:
        logger.error(f"Threads image extraction failed: {str(e)}")
        raise ValueError("No downloadable media was found in this Threads post.")

def get_media_info(url):
    platform = is_supported_url(url)
    if not platform:
        raise ValueError("Unsupported platform or invalid URL.")

    original_url = url
    if platform == 'threads':
        url = url.replace('threads.com', 'threads.net')

    ydl_opts = {
        'quiet': True,
        'no_warnings': True,
        'no_color': True,
        'skip_download': True,
        'geo_bypass': True,
        'extract_flat': 'in_playlist',
        'cookiefile': '/app/secrets/youtube_cookies.txt',
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
    except yt_dlp.utils.DownloadError as e:
        import re
        raw_msg = str(e)
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        clean_msg = ansi_escape.sub('', raw_msg)
        
        logger.error(f"yt-dlp DownloadError [{platform}] exc_class={type(e).__name__} msg={clean_msg!r}")
        
        if platform == "threads":
            if "Unsupported URL" in clean_msg or "not supported" in clean_msg.lower():
                raise ValueError("Unable to access this public Threads post.")
            elif "Sign in" in clean_msg or "Private" in clean_msg or "login" in clean_msg.lower() or "401" in clean_msg or "403" in clean_msg:
                raise ValueError("This Threads post is not publicly accessible.")
            elif "an image post" in clean_msg.lower() or "has no downloadable video" in clean_msg.lower() or "no video post found" in clean_msg.lower():
                import re
                m = re.search(r'Post "([^"]+)"', clean_msg)
                if m:
                    canonical_url = f"https://www.threads.net/t/{m.group(1)}"
                else:
                    canonical_url = original_url
                return extract_threads_images(canonical_url)
            else:
                raise ValueError("No downloadable media was found in this Threads post.")
        elif platform == "youtube":
            if "Sign in" in clean_msg or "Private" in clean_msg or "This video is unavailable" in clean_msg or "members only" in clean_msg.lower():
                raise ValueError("This YouTube video is not publicly accessible.")
            elif "429" in clean_msg or "too many requests" in clean_msg.lower():
                raise ValueError("YouTube rate-limited this request. Please try again in a moment.")
            elif "bot" in clean_msg.lower() or "confirm you're not" in clean_msg.lower():
                raise ValueError("YouTube is requesting a verification challenge on this server. Please try again later.")
            else:
                raise ValueError("Unable to analyze this YouTube video. Please try again later.")
        else:
            raise ValueError("Unable to analyze this media URL.")
    except Exception as e:
        import re
        raw_msg = str(e)
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        clean_msg = ansi_escape.sub('', raw_msg)
        
        logger.error(f"yt-dlp non-DownloadError [{platform}] exc_class={type(e).__name__} repr={repr(e)!r}")
        
        # UnicodeEncodeError means extraction succeeded but yt-dlp failed printing to the console.
        # This is a Windows console encoding issue – not an extraction failure.
        # The info dict was already returned above; getting here means extract_info itself raised.
        if isinstance(e, UnicodeEncodeError):
            raise ValueError("Media extraction encountered a text encoding issue on the server. Please try again.")
        
        raise ValueError("Unable to analyze this media URL. Please try again later.")

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

    original_url = url
    if platform == 'threads':
        url = url.replace('threads.com', 'threads.net')
        if format_id and format_id.startswith('img_'):
            image_info = extract_threads_images(original_url)
            img_url = None
            for img in image_info['image_formats']:
                if img['id'] == format_id:
                    img_url = img['url']
                    break
            
            if not img_url:
                raise ValueError("Image not found in the post.")
                
            temp_dir = tempfile.gettempdir()
            unique_filename = f"dl_{uuid.uuid4().hex}"
            filepath = os.path.join(temp_dir, f"{unique_filename}.jpg")
            
            try:
                req = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'})
                with urllib.request.urlopen(req, timeout=15) as response, open(filepath, 'wb') as out_file:
                    out_file.write(response.read())
                    
                safe_title = sanitize_filename(image_info['title']) or 'threads_image'
                return filepath, f"{safe_title}.jpg"
            except Exception as e:
                logger.error(f"Threads image download failed: {str(e)}")
                raise ValueError("Failed to download Threads image.")

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
