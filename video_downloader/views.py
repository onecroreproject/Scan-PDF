import json
import os
from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.views.decorators.http import require_http_methods
from django.views.decorators.csrf import csrf_exempt
from urllib.parse import urlparse
from . import services

def render_downloader(request, platform_name, meta_title=None, meta_description=None, keywords=None):
    context = {
        'page_title': meta_title or f'{platform_name} Video Downloader - ScanPDF',
        'meta_description': meta_description or f'Free online {platform_name} video downloader. Download videos from {platform_name} in high quality.',
        'platform_name': platform_name,
        'keywords': keywords or f'{platform_name} Video Downloader, Online Video Downloader, Download {platform_name} Video'
    }
    return render(request, 'video_downloader/index.html', context)

def index(request):
    """Renders the main video downloader page."""
    return render_downloader(request, 'Universal', 
        meta_title='Online Video Downloader - ScanPDF',
        meta_description='Free online video downloader. Download videos from YouTube, Facebook, Instagram, TikTok, Twitter and more in high quality.',
        keywords='YouTube Video Downloader, Facebook Video Downloader, Instagram Video Downloader, TikTok Video Downloader, X Video Downloader, Online Video Downloader, Video Downloader'
    )

def youtube_downloader(request):
    return render_downloader(request, 'YouTube',
        meta_title='YouTube Video Downloader - Fast & Free | ScanPDF',
        meta_description='Download YouTube videos easily in MP4 or MP3 format. Free, fast, and secure YouTube video downloader.',
        keywords='YouTube Video Downloader, download youtube videos, save youtube video, youtube to mp4'
    )

def facebook_downloader(request):
    return render_downloader(request, 'Facebook',
        meta_title='Facebook Video Downloader - Fast & Free | ScanPDF',
        meta_description='Download Facebook videos directly to your device. High quality, free, and secure Facebook video downloader.',
        keywords='Facebook Video Downloader, fb video downloader, download facebook video, save facebook video'
    )

def twitter_downloader(request):
    return render_downloader(request, 'X (Twitter)',
        meta_title='X (Twitter) Video Downloader - Fast & Free | ScanPDF',
        meta_description='Download videos and GIFs from X (formerly Twitter). Free online Twitter video downloader.',
        keywords='Twitter Video Downloader, X Video Downloader, download twitter video, save twitter video'
    )

def instagram_downloader(request):
    return render_downloader(request, 'Instagram',
        meta_title='Instagram Video Downloader - Fast & Free | ScanPDF',
        meta_description='Download Instagram Reels, IGTV, and videos. Free online Instagram video downloader.',
        keywords='Instagram Video Downloader, IG video downloader, download instagram reels, save instagram video'
    )

def tiktok_downloader(request):
    return render_downloader(request, 'TikTok',
        meta_title='TikTok Video Downloader - Fast & Free | ScanPDF',
        meta_description='Download TikTok videos without watermark. Fast and free online TikTok video downloader.',
        keywords='TikTok Video Downloader, download tiktok without watermark, save tiktok video'
    )

def vimeo_downloader(request):
    return render_downloader(request, 'Vimeo',
        meta_title='Vimeo Video Downloader - Fast & Free | ScanPDF',
        meta_description='Download Vimeo videos in HD quality. Free and fast online Vimeo video downloader.',
        keywords='Vimeo Video Downloader, download vimeo video, save vimeo video'
    )

def reddit_downloader(request):
    return render_downloader(request, 'Reddit',
        meta_title='Reddit Video Downloader - Fast & Free | ScanPDF',
        meta_description='Download Reddit videos with audio. Fast, free, and secure Reddit video downloader.',
        keywords='Reddit Video Downloader, download reddit video, save reddit video'
    )

def dailymotion_downloader(request):
    return render_downloader(request, 'Dailymotion',
        meta_title='Dailymotion Video Downloader - Fast & Free | ScanPDF',
        meta_description='Download Dailymotion videos in high quality. Free online Dailymotion video downloader.',
        keywords='Dailymotion Video Downloader, download dailymotion video, save dailymotion video'
    )

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
            
        # Analyze using service
        result = services.analyze_video(url)
        
        return JsonResponse(result)
        
    except ValueError as e:
        return JsonResponse({'error': str(e)}, status=400)
    except Exception as e:
        return JsonResponse({'error': f"Server execution error: {str(e)}"}, status=500)

@require_http_methods(["POST", "GET"])
def download_video(request):
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
            
        # Download format
        filepath, title = services.download_format(url, format_id, format_type)
        
        if not filepath or not os.path.exists(filepath):
            return JsonResponse({'error': 'Failed to download file'}, status=500)
            
        # Prepare response
        filename = os.path.basename(filepath)
        _, ext = os.path.splitext(filename)
        
        # Make a safe title
        safe_title = "".join([c for c in title if c.isalpha() or c.isdigit() or c==' ']).rstrip()
        safe_title = safe_title.replace(' ', '_')
        if not safe_title:
            safe_title = 'video'
            
        download_name = f"{safe_title}{ext}"
        
        # Return FileResponse (file will be kept open until fully streamed)
        response = FileResponse(open(filepath, 'rb'), as_attachment=True, filename=download_name)
        return response
        
    except ValueError as e:
        return JsonResponse({'error': str(e)}, status=400)
    except Exception as e:
        return JsonResponse({'error': f"Download failed: {str(e)}"}, status=500)
