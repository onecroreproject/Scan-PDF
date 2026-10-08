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

def get_all_image_candidates(candidates_list):
    if not isinstance(candidates_list, list):
        return []
    valid_cands = []
    for c in candidates_list:
        if isinstance(c, dict):
            url = c.get('url') or c.get('src')
            if url:
                w = c.get('width') or c.get('config_width') or 0
                h = c.get('height') or c.get('config_height') or 0
                valid_cands.append({'url': url, 'width': w, 'height': h})
                
    sorted_cands = sorted(
        valid_cands,
        key=lambda x: (x.get('width', 0) * x.get('height', 0)),
        reverse=True
    )
    return sorted_cands

def get_all_video_candidates(candidates_list):
    if not isinstance(candidates_list, list):
        return []
    valid_cands = [c for c in candidates_list if isinstance(c, dict) and 'url' in c]
    if not valid_cands:
        return []
    
    # Sort from low to high resolution (by height)
    sorted_cands = sorted(
        valid_cands,
        key=lambda x: (x.get('height', 0), x.get('width', 0))
    )
    
    # Deduplicate: if multiple have the same height, prefer the one with '.mp4' in URL
    unique_resolutions = {}
    for c in sorted_cands:
        h = c.get('height', 0)
        # If we don't have this height yet, or if this candidate is an mp4 and the existing one isn't
        if h not in unique_resolutions:
            unique_resolutions[h] = c
        else:
            if '.mp4' in c['url'] and '.mp4' not in unique_resolutions[h]['url']:
                unique_resolutions[h] = c

    return list(unique_resolutions.values())

def extract_media_from_post_object(post_obj):
    images = []
    videos = []
    
    def process_item(item):
        has_vid = False
        if 'video_versions' in item and isinstance(item['video_versions'], list):
            vids = get_all_video_candidates(item['video_versions'])
            if vids:
                videos.extend(vids)
                has_vid = True
        
        if 'image_versions2' in item and 'candidates' in item['image_versions2']:
            cands = get_all_image_candidates(item['image_versions2']['candidates'])
            if cands and not has_vid:
                for c in cands: c['source'] = 'image_versions2'
                images.extend(cands)
        elif 'display_resources' in item and isinstance(item['display_resources'], list):
            cands = get_all_image_candidates(item['display_resources'])
            if cands and not has_vid:
                for c in cands: c['source'] = 'display_resources'
                images.extend(cands)

    if 'carousel_media' in post_obj and isinstance(post_obj['carousel_media'], list):
        for item in post_obj['carousel_media']:
            if isinstance(item, dict):
                process_item(item)
    else:
        process_item(post_obj)
        
    if not images and not videos:
        if 'display_url' in post_obj and isinstance(post_obj['display_url'], str):
            images.append({'url': post_obj['display_url'], 'width': 0, 'height': 0, 'source': 'display_url'})
        elif 'image_url' in post_obj and isinstance(post_obj['image_url'], str):
            images.append({'url': post_obj['image_url'], 'width': 0, 'height': 0, 'source': 'image_url'})
            
    return images, videos

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
        if 'image_versions2' in data or 'carousel_media' in data or 'candidates' in data or 'display_resources' in data:
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
                        if ('"code"' in block or '"shortcode"' in block or '"image_versions2"' in block or '"candidates"' in block or '"display_resources"' in block or '"display_url"' in block):
                            try:
                                obj = json.loads(block)
                                results.append(obj)
                            except:
                                pass
                        break
    return results

def resolve_threads_url(url):
    import urllib.request
    url = url.split('?')[0]
    
    if '/t/' in url or '/share/' in url or '/post/' not in url:
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
            with urllib.request.urlopen(req, timeout=10) as response:
                url = response.geturl().split('?')[0]
        except Exception:
            pass
            
    url = url.replace('threads.net', 'threads.com')
    return url

def extract_threads_media(original_url):
    import re
    canonical_url = resolve_threads_url(original_url)
    
    try:
        opener = urllib.request.build_opener(SafeThreadsRedirectHandler())
        req = urllib.request.Request(
            canonical_url, 
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
        candidate_videos = []
        parsed_objects = []
        
        # 1. Parse all script blocks for JSON robustly
        script_matches = re.findall(r'<script[^>]*>([\s\S]*?)</script>', html_content)
        for s in script_matches:
            if 'image_versions2' in s or 'video_versions' in s or 'candidates' in s or 'display_url' in s or 'carousel_media' in s or 'RelayPreloadedState' in s or 'requireLazy' in s:
                parsed_objects.extend(extract_all_json_from_string(s))

        # 2. Identify the requested post object
        shortcode_match = re.search(r'/(?:t|share|post)/([a-zA-Z0-9_-]+)', canonical_url)
        shortcode = shortcode_match.group(1) if shortcode_match else None
        
        post_obj = None
        for obj in parsed_objects:
            if shortcode:
                post_obj = find_post_object(obj, shortcode)
                if post_obj:
                    break
                    
        # 3. Extract candidate media
        if post_obj:
            imgs, vids = extract_media_from_post_object(post_obj)
            candidate_images.extend(imgs)
            candidate_videos.extend(vids)
        else:
            for obj in parsed_objects:
                any_obj = find_any_media_object(obj)
                if any_obj:
                    imgs, vids = extract_media_from_post_object(any_obj)
                    candidate_images.extend(imgs)
                    candidate_videos.extend(vids)
                    if candidate_images or candidate_videos:
                        break
                        
        # 4. Fallback to generic JSON-LD and Meta tags if needed
        if not candidate_images and not candidate_videos:
            for obj in parsed_objects:
                if isinstance(obj, dict):
                    if obj.get('@type') == 'ImageObject':
                        if 'contentUrl' in obj: candidate_images.append(obj['contentUrl'])
                        elif 'url' in obj: candidate_images.append(obj['url'])
                    elif obj.get('@type') == 'VideoObject':
                        if 'contentUrl' in obj: candidate_videos.append({'url': obj['contentUrl']})
                        elif 'url' in obj: candidate_videos.append({'url': obj['url']})
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
                elif 'property="og:video"' in tag or 'property="og:video:url"' in tag:
                    content_match = re.search(r'content="([^"]+)"', tag)
                    if content_match:
                        val = html.unescape(content_match.group(1))
                        candidate_videos.append({'url': val})
                        
        # 5. Deduplicate and validate images
        unique_images = []
        for img in candidate_images:
            if isinstance(img, dict):
                img = img.get('url')
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
                    
        # 6. Deduplicate videos
        unique_videos = []
        seen_vid_urls = set()
        for vid in candidate_videos:
            url = vid.get('url')
            if not url or not isinstance(url, str):
                continue
            url = url.replace('\\/', '/')
            if url not in seen_vid_urls and url.startswith('http'):
                vid['url'] = url
                unique_videos.append(vid)
                seen_vid_urls.add(url)
                
        # Diagnostic logging
        script_types = list(set(re.findall(r'<script[^>]*type="([^"]+)"', html_content)))
        media_key_hits = []
        if 'image_versions2' in html_content: media_key_hits.append('image_versions2')
        if 'video_versions' in html_content: media_key_hits.append('video_versions')
        if 'candidates' in html_content: media_key_hits.append('candidates')
        if 'carousel_media' in html_content: media_key_hits.append('carousel_media')
        
        logger.info(f"Threads fallback: shortcode={shortcode} status={status} videos={len(unique_videos)} images={len(unique_images)} method=hydration media_keys={media_key_hits}")
                
        if not unique_images and not unique_videos:
            logger.error(f"Threads media fallback failed: original_host={urllib.parse.urlparse(original_url).hostname} final_host={urllib.parse.urlparse(final_url).hostname} status={status} html_length={len(html_content)} script_count={len(script_matches)} requested_obj_found={post_obj is not None} media_keys={media_key_hits} candidate_images={len(candidate_images)} candidate_videos={len(candidate_videos)}")
            raise ValueError("No downloadable media was found in this Threads post.")
        
        title_match = re.search(r'<meta[^>]+property="og:title"[^>]+content="([^"]+)"', html_content)
        if not title_match:
            title_match = re.search(r'<meta[^>]+content="([^"]+)"[^>]+property="og:title"', html_content)
        title = html.unescape(title_match.group(1)) if title_match else "Threads Media"
        
        thumbnail = unique_images[0] if unique_images else None
        if not thumbnail and unique_videos and 'url' in unique_videos[0]:
            # Fallback thumbnail logic could go here, but video URL itself isn't a thumbnail
            thumbnail = ""
            
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
            
        video_items = []
        for idx, vid in enumerate(unique_videos):
            h = vid.get('height') or 0
            w = vid.get('width') or 0
            quality = f"{h}p" if h else "Best"
            
            video_items.append({
                'id': f'vid_{idx}',
                'type': 'video',
                'quality': quality,
                'format': 'mp4',
                'width': w,
                'height': h,
                'fileSize': None,
                'has_audio': True,
                'can_download_mp3': True,
                'url': vid['url']
            })
            
        # Sort video items by height descending
        video_items.sort(key=lambda x: x.get('height') or 0, reverse=True)
            
        has_any_audio = len(video_items) > 0
        audio_info = {
            'available': has_any_audio,
            'qualities': [
                {'label': '320 kbps', 'value': '320'},
                {'label': '256 kbps', 'value': '256'},
                {'label': '192 kbps', 'value': '192'},
                {'label': '128 kbps', 'value': '128'},
                {'label': '96 kbps', 'value': '96'},
                {'label': '64 kbps', 'value': '64'},
            ] if has_any_audio else []
        }
            
        return {
            'platform': 'threads',
            'title': title,
            'author': 'Threads',
            'thumbnail': thumbnail,
            'duration': None,
            'video_formats': video_items,
            'audio': audio_info,
            'image_formats': image_items
        }
    except Exception as e:
        logger.error(f"Threads media extraction failed: {str(e)}")
        raise ValueError("No downloadable media was found in this Threads post.")

def extract_instagram_media(original_url):
    import urllib.request, re, json, urllib.parse, html
    parsed = urllib.parse.urlparse(original_url)
    shortcode_match = re.search(r'/(?:p|reel|tv)/([a-zA-Z0-9_-]+)', parsed.path)
    if not shortcode_match:
        raise ValueError("Invalid Instagram URL")
    shortcode = shortcode_match.group(1)

    embed_url = f"https://www.instagram.com/p/{shortcode}/embed/captioned/"
    req = urllib.request.Request(embed_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'})
    
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            html_content = resp.read().decode('utf-8', errors='ignore')
    except Exception as e:
        logger.error(f"Insta embed fallback failed: {str(e)}")
        html_content = ""
        
    candidate_images = []
    
    def extract_from_html(html_str):
        script_matches = re.findall(r'<script[^>]*>([\s\S]*?)</script>', html_str)
        parsed_objects = []
        for s in script_matches:
            if 'image_versions2' in s or 'carousel_media' in s or 'display_url' in s or 'shortcode_media' in s or 'display_resources' in s:
                parsed_objects.extend(extract_all_json_from_string(s))
                
        post_obj = None
        for obj in parsed_objects:
            post_obj = find_post_object(obj, shortcode)
            if post_obj:
                break
                
        extracted_imgs = []
        if post_obj:
            imgs, vids = extract_media_from_post_object(post_obj)
            extracted_imgs.extend(imgs)
        else:
            for obj in parsed_objects:
                if isinstance(obj, dict) and 'shortcode_media' in obj:
                    imgs, vids = extract_media_from_post_object(obj['shortcode_media'])
                    extracted_imgs.extend(imgs)
                    if extracted_imgs:
                        break
                        
        # Robust regex parsing
        for m in re.finditer(r'\{"src":"([^"]+)","config_width":(\d+),"config_height":(\d+)\}', html_str):
            url = m.group(1).replace('\\/', '/').replace('\\u0026', '&').replace('\\\\', '\\')
            if 'profile' not in url.lower():
                extracted_imgs.append({'url': url, 'width': int(m.group(2)), 'height': int(m.group(3)), 'source': 'display_resources_regex'})
                
        for m in re.finditer(r'\{"url":"([^"]+)","width":(\d+),"height":(\d+)\}', html_str):
            url = m.group(1).replace('\\/', '/').replace('\\u0026', '&').replace('\\\\', '\\')
            if '.mp4' not in url and 'profile' not in url.lower():
                extracted_imgs.append({'url': url, 'width': int(m.group(2)), 'height': int(m.group(3)), 'source': 'image_versions2_regex'})
                
        return extracted_imgs

    # 1. Try JSON from embed HTML
    candidate_images.extend(extract_from_html(html_content))
    
    # 2. Try JSON and meta tags from regular page HTML
    if not candidate_images:
        req_reg = urllib.request.Request(
            f"https://www.instagram.com/p/{shortcode}/", 
            headers={
                'User-Agent': 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
            }
        )
        try:
            with urllib.request.urlopen(req_reg, timeout=10) as resp_reg:
                html_reg = resp_reg.read().decode('utf-8', errors='ignore')
                
            candidate_images.extend(extract_from_html(html_reg))
            
            # 3. Try strict og:image from regular page
            if not candidate_images:
                og_url_match = re.search(r'<meta[^>]+property="og:url"[^>]+content="([^"]+)"', html_reg)
                if og_url_match and shortcode in og_url_match.group(1):
                    og_image_match = re.search(r'<meta[^>]+property="og:image"[^>]+content="([^"]+)"', html_reg)
                    if og_image_match:
                        candidate_images.append({'url': html.unescape(og_image_match.group(1)), 'width': 0, 'height': 0, 'source': 'og:image'})
        except Exception as e:
            logger.error(f"Insta regular page fallback failed: {str(e)}")

    # 4. Try strict EmbeddedMediaImage from embed HTML
    if not candidate_images and html_content:
        for m in re.finditer(r'<img class="EmbeddedMediaImage"[^>]+src="([^"]+)"', html_content):
            candidate_images.append({'url': html.unescape(m.group(1)), 'width': 640, 'height': 640, 'source': 'EmbeddedMediaImage'})

    logger.info(f"Instagram target shortcode: {shortcode}")
    unique_images = []
    seen_urls = set()
    for u in candidate_images:
        if isinstance(u, str):
            u = {'url': u, 'width': 0, 'height': 0, 'source': 'unknown'}
        url_str = u['url'].replace('\\/', '/')
        if url_str not in seen_urls and url_str.startswith('http') and 'profile' not in url_str.lower():
            u['url'] = url_str
            unique_images.append(u)
            seen_urls.add(url_str)
            
    unique_images.sort(key=lambda x: x.get('width', 0) * x.get('height', 0), reverse=True)

    for c in unique_images:
        logger.info(f"candidate: {c['width']}x{c['height']} source={c.get('source', 'unknown')}")

    if not unique_images:
        logger.error(f"Instagram image extraction failed: shortcode={shortcode} no target images found.")
        raise ValueError("Unable to analyze this media URL. Target post media could not be located.")

    selected = unique_images[0]
    logger.info(f"SELECTED: {selected['width']}x{selected['height']} source={selected.get('source', 'unknown')}")

    image_formats = []
    # Provide the single best selected image
    image_formats.append({
        'id': 'img_0',
        'type': 'image',
        'quality': f"{selected['width']}x{selected['height']}" if selected['width'] else 'Best available',
        'format': 'jpg',
        'fileSize': None,
        'has_audio': False,
        'can_download_mp3': False,
        'url': selected['url']
    })

    return {
        'platform': 'instagram',
        'title': f'Instagram Post {shortcode}',
        'author': 'Instagram',
        'thumbnail': selected['url'],
        'duration': None,
        'video_formats': [],
        'audio': {'available': False, 'qualities': []},
        'image_formats': image_formats
    }

def extract_twitter_media(original_url):
    import urllib.request, re, json, urllib.parse, html
    parsed = urllib.parse.urlparse(original_url)
    match = re.search(r'status/(\d+)', parsed.path)
    if not match:
        raise ValueError("Invalid Twitter URL")
    tweet_id = match.group(1)
    
    # 1. Fetch the existing public X/Twitter page using Googlebot
    req = urllib.request.Request(original_url, headers={
        'User-Agent': 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    })
    
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            html_content = resp.read().decode('utf-8', errors='ignore')
    except Exception as e:
        logger.error(f"Twitter public page fetch failed: {str(e)}")
        html_content = ""
        
    candidate_images = []
    
    if html_content:
        # Extract from meta tags
        for m in re.finditer(r'<meta[^>]+property="(?:og:image|twitter:image)"[^>]+content="([^"]+)"', html_content):
            img_url = html.unescape(m.group(1))
            if 'profile' not in img_url and 'twimg.com/' in img_url:
                candidate_images.append(img_url)

        # Extract ALL pbs.twimg.com/media/ links from the HTML (covers JSON payload in SSR)
        for m in re.finditer(r'(https://pbs\.twimg\.com/media/[a-zA-Z0-9_-]+(?:(?:\?format=|\.)(?:jpg|png|webp))[^"\'\s\\]*)', html_content):
            img_url = m.group(1).replace('\\u0026', '&')
            candidate_images.append(img_url)

    # 2. If no images found, fallback to Syndication API
    if not candidate_images:
        logger.info("Twitter Googlebot SSR returned no images, trying syndication fallback")
        syn_url = f"https://cdn.syndication.twimg.com/tweet-result?id={tweet_id}"
        req2 = urllib.request.Request(syn_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'})
        try:
            with urllib.request.urlopen(req2, timeout=10) as resp2:
                data = json.loads(resp2.read().decode('utf-8'))
                if 'photos' in data:
                    for p in data['photos']:
                        candidate_images.append(p['url'])
        except Exception as e:
            logger.error(f"Twitter syndication fallback failed: {str(e)}")

    # 3. If still no images, fallback to vxtwitter API
    if not candidate_images:
        logger.info("Twitter Syndication failed, trying vxtwitter API fallback")
        vx_url = f"https://api.vxtwitter.com/Twitter/status/{tweet_id}"
        req3 = urllib.request.Request(vx_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'})
        try:
            with urllib.request.urlopen(req3, timeout=10) as resp3:
                data = json.loads(resp3.read().decode('utf-8'))
                if 'media_extended' in data:
                    for media in data['media_extended']:
                        if media.get('type') == 'image' and 'url' in media:
                            candidate_images.append(media['url'])
                elif 'mediaURLs' in data:
                    candidate_images.extend(data['mediaURLs'])
        except Exception as e:
            logger.error(f"Twitter vxtwitter API fallback failed: {str(e)}")

    # 4. If still no images, fallback to fxtwitter API
    if not candidate_images:
        logger.info("Twitter vxtwitter API failed, trying fxtwitter API fallback")
        fx_url = f"https://api.fxtwitter.com/status/{tweet_id}"
        req4 = urllib.request.Request(fx_url, headers={'User-Agent': 'Mozilla/5.0'})
        try:
            with urllib.request.urlopen(req4, timeout=10) as resp4:
                data = json.loads(resp4.read().decode('utf-8'))
                if 'tweet' in data and 'media' in data['tweet'] and 'photos' in data['tweet']['media']:
                    for p in data['tweet']['media']['photos']:
                        if 'url' in p:
                            candidate_images.append(p['url'])
        except Exception as e:
            logger.error(f"Twitter fxtwitter API fallback failed: {str(e)}")

    # Deduplicate and upgrade quality
    unique_images = []
    for u in candidate_images:
        if 'profile' in u or 'video_thumb' in u:
            continue
            
        # Upgrade quality
        u = re.sub(r'name=[^&]+', 'name=large', u)
        if ':large' not in u and not re.search(r'name=(?:orig|large)', u):
            if '?' not in u:
                u = u + ':large'
        
        # Deduplicate using base ID
        base_match = re.search(r'media/([a-zA-Z0-9_-]+)', u)
        base_id = base_match.group(1) if base_match else u
        
        if not any(base_id in existing for existing in unique_images):
            unique_images.append(u)

    if not unique_images:
        raise ValueError("No downloadable images could be found in this X/Twitter post.")
        
    image_formats = []
    for idx, img in enumerate(unique_images):
        image_formats.append({
            'id': f'img_{idx}',
            'type': 'image',
            'quality': 'Best available',
            'format': 'jpg' if 'png' not in img else 'png',
            'fileSize': None,
            'has_audio': False,
            'can_download_mp3': False,
            'url': img
        })
        
    title = f'X Post {tweet_id}'
    author = 'X User'
    if html_content:
        title_match = re.search(r'<meta[^>]+property="og:description"[^>]+content="([^"]+)"', html_content)
        if title_match: title = html.unescape(title_match.group(1))
        
        author_match = re.search(r'<meta[^>]+property="og:title"[^>]+content="([^"]+)"', html_content)
        if author_match: author = html.unescape(author_match.group(1))
    
    logger.info(f"Twitter image fallback: status_id={tweet_id} yt_dlp_video=false image_candidates={len(candidate_images)} validated_images={len(unique_images)} method=googlebot_ssr result=success")

    return {
        'platform': 'twitter',
        'title': title,
        'author': author,
        'thumbnail': unique_images[0],
        'duration': None,
        'video_formats': [],
        'audio': {'available': False, 'qualities': []},
        'image_formats': image_formats
    }

def extract_facebook_media(original_url):
    import urllib.request, re, urllib.parse, html
    
    resolved_url = original_url
    if '/share/' in original_url:
        req = urllib.request.Request(original_url, headers={'User-Agent': 'Mozilla/5.0'})
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                resolved_url = resp.geturl()
        except:
            pass
            
    req = urllib.request.Request(resolved_url, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'en-US,en;q=0.9'
    })
    
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            html_content = resp.read().decode('utf-8', errors='ignore')
    except Exception:
        raise ValueError("Unable to analyze this media URL.")

    candidate_images = []
    
    for match in re.finditer(r'<meta[^>]+property="og:image"[^>]+content="([^"]+)"', html_content):
        val = html.unescape(match.group(1))
        if 'fbcdn.net' in val and 'profile' not in val:
            candidate_images.append(val)
            
    script_matches = re.findall(r'<script[^>]*>([\s\S]*?)</script>', html_content)
    for s in script_matches:
        if 'image_uri' in s or 'imageURL' in s or 'browser_native_hd_url' in s:
            urls = re.findall(r'"(https://[^"]+fbcdn\.net[^"]+)"', s)
            urls = [u.replace('\\/', '/') for u in urls]
            for u in urls:
                if 'profile' not in u and ('_nc_cat' in u or 'fbst' in u):
                    candidate_images.append(u)

    unique_images = []
    for u in candidate_images:
        if u not in unique_images:
            unique_images.append(u)
            
    if not unique_images:
        post_id = re.search(r'/(?:posts|photo|p)/([^/?]+)', resolved_url)
        logger.error(f"Facebook image fallback failed: post_id={post_id.group(1) if post_id else 'unknown'}")
        raise ValueError("Unable to analyze this media URL.")

    parsed_host = urllib.parse.urlparse(resolved_url).hostname
    logger.info(f"Facebook image fallback: resolved_host={parsed_host} candidates={len(unique_images)} validated=True")

    image_formats = []
    for idx, img in enumerate(unique_images):
        image_formats.append({
            'id': f'img_{idx}',
            'type': 'image',
            'quality': 'Best available',
            'format': 'jpg',
            'fileSize': None,
            'has_audio': False,
            'can_download_mp3': False,
            'url': img
        })

    title_match = re.search(r'<title>(.*?)</title>', html_content)
    title = html.unescape(title_match.group(1)) if title_match else 'Facebook Post'

    return {
        'platform': 'facebook',
        'title': title,
        'author': 'Facebook',
        'thumbnail': unique_images[0],
        'duration': None,
        'video_formats': [],
        'audio': {'available': False, 'qualities': []},
        'image_formats': image_formats
    }

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
    }
    
    if platform == 'youtube':
        cookie_path = '/app/secrets/youtube_cookies.txt'
        if os.path.exists(cookie_path):
            ydl_opts['cookiefile'] = cookie_path

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
                        
                        h = f.get('height')
                        if h:
                            quality = f"{h}p"
                        else:
                            quality = f.get('resolution') or f.get('format_note') or "Best"
                        
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
                            if item.get('has_audio') and not curr.get('has_audio'):
                                unique_video_items[q] = item
                            elif item.get('has_audio') == curr.get('has_audio'):
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
                    {'label': '320 kbps', 'value': '320'},
                    {'label': '256 kbps', 'value': '256'},
                    {'label': '192 kbps', 'value': '192'},
                    {'label': '128 kbps', 'value': '128'},
                    {'label': '96 kbps', 'value': '96'},
                    {'label': '64 kbps', 'value': '64'},
                ] if has_any_audio else []
            }

            if not video_formats and not image_formats:
                if platform == 'instagram':
                    return extract_instagram_media(original_url)
                elif platform == 'twitter':
                    return extract_twitter_media(original_url)
                elif platform == 'facebook':
                    return extract_facebook_media(original_url)
                    
            if platform == 'instagram' and image_items and not video_formats:
                thumbnail = image_items[0].get('url', thumbnail)

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
        
        if platform == "instagram":
            if "no video" in clean_msg.lower():
                return extract_instagram_media(original_url)
            else:
                # If it's a generic yt-dlp error, raise ValueError so it returns a clean 400 instead of 500
                raise ValueError("Unable to analyze this media URL. " + clean_msg.split(':', 1)[-1].strip())
        elif platform == "twitter":
            return extract_twitter_media(original_url)
        elif platform == "facebook":
            return extract_facebook_media(original_url)
        elif platform == "threads":
            if "Unsupported URL" in clean_msg or "not supported" in clean_msg.lower():
                # yt-dlp doesn't support this URL (e.g., /share/ or general threads). Use fallback.
                return extract_threads_media(original_url)
            elif "Sign in" in clean_msg or "Private" in clean_msg or "login" in clean_msg.lower() or "401" in clean_msg or "403" in clean_msg:
                raise ValueError("This Threads post is not publicly accessible.")
            elif "an image post" in clean_msg.lower() or "has no downloadable video" in clean_msg.lower() or "no video post found" in clean_msg.lower():
                m = re.search(r'Post "([^"]+)"', clean_msg)
                if m:
                    canonical_url = f"https://www.threads.net/t/{m.group(1)}"
                else:
                    canonical_url = original_url
                return extract_threads_media(canonical_url)
            else:
                return extract_threads_media(original_url)
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
        import traceback
        full_traceback = traceback.format_exc()
        
        import re
        raw_msg = str(e)
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        clean_msg = ansi_escape.sub('', raw_msg)
        
        logger.error(f"yt-dlp non-DownloadError [{platform}] exc_class={type(e).__name__} repr={repr(e)!r}\nTraceback:\n{full_traceback}")
        
        if platform == "threads" and isinstance(e, FileNotFoundError):
            return extract_threads_media(original_url)
            
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
        if format_id and (format_id.startswith('img_') or format_id.startswith('vid_')):
            media_info = extract_threads_media(original_url)
            
            target_url = None
            if format_id.startswith('img_'):
                for img in media_info['image_formats']:
                    if img['id'] == format_id:
                        target_url = img['url']
                        break
                ext = 'jpg'
            else:
                for vid in media_info['video_formats']:
                    if vid['id'] == format_id:
                        target_url = vid['url']
                        break
                ext = 'mp4'
            
            if not target_url:
                raise ValueError("Media not found in the post.")
                
            temp_dir = tempfile.gettempdir()
            unique_filename = f"dl_{uuid.uuid4().hex}"
            filepath = os.path.join(temp_dir, f"{unique_filename}.{ext}")
            
            try:
                req = urllib.request.Request(target_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
                with urllib.request.urlopen(req, timeout=30) as response, open(filepath, 'wb') as out_file:
                    out_file.write(response.read())
                    
                safe_title = sanitize_filename(media_info['title']) or 'threads_media'
                
                if download_type == "audio" and ext == "mp4":
                    if not is_ffmpeg_available():
                        raise ValueError("MP3 conversion requires FFmpeg on the server.")
                    
                    ffmpeg_loc = get_ffmpeg_location()
                    ffmpeg_cmd = os.path.join(ffmpeg_loc, "ffmpeg") if ffmpeg_loc else "ffmpeg"
                    mp3_filepath = os.path.join(temp_dir, f"{unique_filename}.mp3")
                    
                    import subprocess
                    cmd = [
                        ffmpeg_cmd, "-i", filepath,
                        "-vn", "-b:a", f"{audio_quality}k", "-y", mp3_filepath
                    ]
                    try:
                        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                        os.remove(filepath)
                        return mp3_filepath, f"{safe_title}.mp3"
                    except subprocess.CalledProcessError:
                        logger.error("Failed to convert Threads video to MP3")
                        raise ValueError("Failed to extract audio from Threads video.")
                        
                return filepath, f"{safe_title}.{ext}"
            except Exception as e:
                logger.error(f"Threads media download failed: {str(e)}")
                raise ValueError("Failed to download Threads media.")

    if platform in ['instagram', 'facebook', 'twitter']:
        if format_id and format_id.startswith('img_'):
            # Fetch the info using our fallback
            if platform == 'instagram':
                media_info = extract_instagram_media(original_url)
            elif platform == 'facebook':
                media_info = extract_facebook_media(original_url)
            elif platform == 'twitter':
                media_info = extract_twitter_media(original_url)
            
            target_url = None
            for img in media_info.get('image_formats', []):
                if img['id'] == format_id:
                    target_url = img['url']
                    break
            
            if target_url:
                temp_dir = tempfile.gettempdir()
                unique_filename = f"dl_{uuid.uuid4().hex}"
                ext = 'jpg'
                filepath = os.path.join(temp_dir, f"{unique_filename}.{ext}")
                
                try:
                    req = urllib.request.Request(target_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
                    with urllib.request.urlopen(req, timeout=30) as response, open(filepath, 'wb') as out_file:
                        out_file.write(response.read())
                    safe_title = sanitize_filename(media_info['title']) or f'{platform}_image'
                    return filepath, f"{safe_title}_{format_id}.{ext}"
                except Exception as e:
                    logger.error(f"{platform} image download failed: {str(e)}")
                    # Let it fall through to yt-dlp if it fails, or just raise
                    pass

    temp_dir = tempfile.gettempdir()
    unique_filename = f"dl_{uuid.uuid4().hex}"
    out_tmpl = os.path.join(temp_dir, f"{unique_filename}.%(ext)s")

    ydl_opts = {
        'quiet': True,
        'no_warnings': True,
        'outtmpl': out_tmpl,
    }

     # Use authenticated YouTube cookies only for YouTube downloads
    if platform == 'youtube':
        cookie_path = '/app/secrets/youtube_cookies.txt'
        if os.path.exists(cookie_path):
            ydl_opts['cookiefile'] = cookie_path
    
    ffmpeg_loc = get_ffmpeg_location()

    if download_type == 'audio':
        logger.info(f"YOUTUBE MP3 DIAGNOSTICS: Starting yt-dlp audio fetch. Selected bitrate: {audio_quality} kbps")
        if not is_ffmpeg_available():
            logger.error("YOUTUBE MP3 DIAGNOSTICS: ffmpeg not found.")
            raise ValueError("MP3 conversion requires FFmpeg on the server.")
            
        ydl_opts['format'] = 'bestaudio/best'
        logger.info(f"YOUTUBE MP3 DIAGNOSTICS: Selected format/audio source: {ydl_opts['format']}")
        ydl_opts['postprocessors'] = [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': str(audio_quality),
        }]
        
        if ffmpeg_loc:
            ydl_opts['ffmpeg_location'] = ffmpeg_loc
            logger.info("YOUTUBE MP3 DIAGNOSTICS: ffmpeg_location configured for yt-dlp.")
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
                logger.error("YOUTUBE MP3 DIAGNOSTICS: Output filepath does not exist.")
                raise ValueError("Downloaded file not found. If extracting MP3, ensure FFmpeg is installed.")
                
            title = info.get('title', 'media')
            safe_title = sanitize_filename(title) or 'downloaded_media'
            
            if download_type == 'audio':
                size = os.path.getsize(filepath)
                logger.info(f"YOUTUBE MP3 DIAGNOSTICS: yt-dlp/ffmpeg succeeded. Output file exists: {filepath}, Size: {size} bytes")
                
            return filepath, f"{safe_title}.{ext}"
    except yt_dlp.utils.PostProcessingError as e:
        logger.error(f"yt-dlp postprocessing error: {str(e)}")
        logger.error(f"YOUTUBE MP3 DIAGNOSTICS: ffmpeg postprocessing failure: {e.__class__.__name__} - {str(e)}")
        raise ValueError("MP3 conversion requires FFmpeg on the server.")
    except Exception as e:
        logger.error(f"YOUTUBE MP3 DIAGNOSTICS: Exception during yt-dlp execution: {e.__class__.__name__} - {str(e)}")
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
