import json
import os
from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.views.decorators.http import require_http_methods
from django.views.decorators.csrf import csrf_exempt
from urllib.parse import urlparse
from . import services

def index(request):
    """Renders the main video downloader page."""
    context = {
        'page_title': 'Online Video Downloader - ScanPDF',
        'meta_description': 'Free online video downloader. Download videos from YouTube, Facebook, Instagram, Twitter and more in high quality.',
        'keywords': 'YouTube Video Downloader, Facebook Video Downloader, Instagram Video Downloader, X Video Downloader, Online Video Downloader, Video Downloader'
    }
    return render(request, 'video_downloader/universal.html', context)

def youtube_downloader(request):
    context = {
        'page_title': 'YouTube Video Downloader - Fast & Free | ScanPDF',
        'meta_description': 'Download YouTube videos easily in MP4 or MP3 format. Free, fast, and secure YouTube video downloader.',
        'keywords': 'YouTube Video Downloader, download youtube videos, save youtube video, youtube to mp4'
    }
    return render(request, 'video_downloader/youtube.html', context)

def facebook_downloader(request):
    context = {
        'page_title': 'Facebook Video Downloader - Fast & Free | ScanPDF',
        'meta_description': 'Download Facebook videos directly to your device. High quality, free, and secure Facebook video downloader.',
        'keywords': 'Facebook Video Downloader, fb video downloader, download facebook video, save facebook video'
    }
    return render(request, 'video_downloader/facebook.html', context)

def twitter_downloader(request):
    context = {
        'page_title': 'X (Twitter) Video Downloader - Fast & Free | ScanPDF',
        'meta_description': 'Download videos and GIFs from X (formerly Twitter). Free online Twitter video downloader.',
        'keywords': 'Twitter Video Downloader, X Video Downloader, download twitter video, save twitter video'
    }
    return render(request, 'video_downloader/twitter.html', context)

def instagram_downloader(request):
    context = {
        'page_title': 'Instagram Video Downloader - Fast & Free | ScanPDF',
        'meta_description': 'Download Instagram Reels, IGTV, and videos. Free online Instagram video downloader.',
        'keywords': 'Instagram Video Downloader, IG video downloader, download instagram reels, save instagram video'
    }
    return render(request, 'video_downloader/instagram.html', context)

def threads_downloader(request):
    context = {
        'page_title': 'Threads Video Downloader - Fast & Free | ScanPDF',
        'meta_description': 'Download Threads videos and images. Free online Threads media downloader.',
        'keywords': 'Threads Video Downloader, download threads video, save threads media'
    }
    return render(request, 'video_downloader/threads.html', context)

@csrf_exempt
@require_http_methods(["POST"])
def analyze_url(request):
    """Analyzes the given URL and returns format metadata."""
    try:
        data = json.loads(request.body)
        url = data.get('url')
        
        if not url:
            return JsonResponse({'error': 'URL is required'}, status=400)
            
        # Basic validation
        parsed_url = urlparse(url)
        if not parsed_url.scheme or not parsed_url.netloc:
            return JsonResponse({'error': 'Invalid URL format'}, status=400)
            
        allowed_domains = [
            'youtube.com', 'youtu.be',
            'facebook.com', 'fb.watch',
            'instagram.com',
            'twitter.com', 'x.com',
            'threads.net', 'threads.com'
        ]
        
        hostname = parsed_url.hostname.lower() if parsed_url.hostname else ''
        if not any(hostname == domain or hostname.endswith('.' + domain) for domain in allowed_domains):
            return JsonResponse({'error': 'Unsupported video provider URL'}, status=400)
            
        # Normalize YouTube shorts URL
        if 'youtube.com' in hostname and '/shorts/' in parsed_url.path:
            video_id = parsed_url.path.split('/shorts/')[-1].split('/')[0]
            url = f"https://www.youtube.com/watch?v={video_id}"
            
        # Analyze using service
        result = services.analyze_video(url)
        
        return JsonResponse(result)
        
    except services.YTDLPError as e:
        response_data = {'success': False, 'error_code': e.code, 'message': str(e)}
        from django.conf import settings
        if getattr(settings, 'DEBUG', False) and getattr(e, 'raw_error', None):
            response_data['raw_error'] = e.raw_error
        return JsonResponse(response_data, status=400)
    except ValueError as e:
        return JsonResponse({'success': False, 'error_code': 'VALIDATION_ERROR', 'message': str(e)}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error_code': 'INTERNAL_ERROR', 'message': "An internal error occurred. Please try again."}, status=500)

@require_http_methods(["POST", "GET"])
def download_media(request):
    """Triggers the download for a specific format and returns the file."""
    try:
        # We can handle both GET and POST for download
        # Usually GET with query params is easier for direct browser download
        if request.method == "POST":
            url = request.POST.get('url')
            format_id = request.POST.get('format_id')
            format_type = request.POST.get('format_type')
        else:
            url = request.GET.get('url')
            format_id = request.GET.get('format_id')
            format_type = request.GET.get('format_type')
            
        if not all([url, format_id, format_type]):
            return JsonResponse({'error': 'Missing required parameters'}, status=400)
            
        parsed_url = urlparse(url)
        allowed_domains = [
            'youtube.com', 'youtu.be',
            'facebook.com', 'fb.watch',
            'instagram.com',
            'twitter.com', 'x.com',
            'threads.net', 'threads.com'
        ]
        
        hostname = parsed_url.hostname.lower() if parsed_url.hostname else ''
        if not any(hostname == domain or hostname.endswith('.' + domain) for domain in allowed_domains):
            return JsonResponse({'error': 'Unsupported video provider URL'}, status=400)
            
        # Normalize YouTube shorts URL
        if 'youtube.com' in hostname and '/shorts/' in parsed_url.path:
            video_id = parsed_url.path.split('/shorts/')[-1].split('/')[0]
            url = f"https://www.youtube.com/watch?v={video_id}"
            
        # Check for download_id to set cookie and track progress
        download_id = request.GET.get('download_id')
        if not download_id and request.method == "POST":
            download_id = request.POST.get('download_id')
            
        # Download format
        filepath, title = services.download_format(url, format_id, format_type, download_id=download_id)
        
        if not filepath or not os.path.exists(filepath):
            return JsonResponse({'error': 'Failed to download file'}, status=500)
            
        if os.path.getsize(filepath) == 0:
            os.remove(filepath)
            return JsonResponse({'error': 'Downloaded file is empty'}, status=500)
            
        # Prepare response
        filename = os.path.basename(filepath)
        _, ext = os.path.splitext(filename)
        
        # Make a safe title
        safe_title = "".join([c for c in title if c.isalpha() or c.isdigit() or c==' ']).rstrip()
        safe_title = safe_title.replace(' ', '_')
        if not safe_title:
            safe_title = 'video'
            
        download_name = f"{safe_title}{ext}"
        
        # Ensure correct MIME type for MP3
        kwargs = {'as_attachment': True, 'filename': download_name}
        if ext.lower() == '.mp3':
            kwargs['content_type'] = 'audio/mpeg'
            
        # Return FileResponse (file will be kept open until fully streamed)
        response = FileResponse(open(filepath, 'rb'), **kwargs)
        
        # Check for download_id to set cookie (tells frontend download has started)
        download_id = request.GET.get('download_id')
        if not download_id and request.method == "POST":
            download_id = request.POST.get('download_id')
            
        if download_id:
            response.set_cookie('download_started', download_id, max_age=60, samesite='Lax')
            
        return response
        
    except services.YTDLPError as e:
        response_data = {'success': False, 'error_code': e.code, 'message': str(e)}
        from django.conf import settings
        if getattr(settings, 'DEBUG', False) and getattr(e, 'raw_error', None):
            response_data['raw_error'] = e.raw_error
        return JsonResponse(response_data, status=400)
    except ValueError as e:
        return JsonResponse({'success': False, 'error_code': 'VALIDATION_ERROR', 'message': str(e)}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error_code': 'INTERNAL_ERROR', 'message': "An internal error occurred. Please try again."}, status=500)

@require_http_methods(["GET"])
def download_progress(request):
    """Returns the current download progress for a specific download_id"""
    download_id = request.GET.get('download_id')
    if not download_id:
        return JsonResponse({'error': 'Missing download_id'}, status=400)
        
    from django.core.cache import cache
    progress_data = cache.get(f"dl_prog_{download_id}")
    
    if progress_data:
        return JsonResponse(progress_data)
        
    # If not found in cache, it might be initializing or already finished
    return JsonResponse({'status': 'unknown', 'percent': 0})
