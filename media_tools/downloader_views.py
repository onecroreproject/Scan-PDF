import json
import os
import logging
from django.shortcuts import render
from django.http import JsonResponse, StreamingHttpResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from media_tools.services.downloader_service import get_media_info, download_media_to_temp

logger = logging.getLogger(__name__)

PLATFORMS = [
    {"id": "instagram", "name": "Instagram"},
    {"id": "twitter", "name": "X / Twitter"},
    {"id": "facebook", "name": "Facebook"},
    {"id": "threads", "name": "Threads"},
    {"id": "youtube", "name": "YouTube"},
]

def downloader_page(request, platform=None):
    """Serve the main Video Downloader page."""
    context = {
        "platforms": PLATFORMS,
        "selected_platform": platform or "",
        "seo_title": f"{platform.capitalize() if platform else 'Video'} Downloader - ScanPDF",
        "seo_description": "Download supported public video and image media from supported platforms.",
    }
    return render(request, "media_tools/downloader.html", context)

def instagram_downloader(request):
    return render(request, "media_tools/instagram_downloader.html", {"platform": "instagram"})

def twitter_downloader(request):
    return render(request, "media_tools/twitter_downloader.html", {"platform": "twitter"})

def facebook_downloader(request):
    return render(request, "media_tools/facebook_downloader.html", {"platform": "facebook"})

def threads_downloader(request):
    return render(request, "media_tools/threads_downloader.html", {"platform": "threads"})

def youtube_downloader(request):
    return render(request, "media_tools/youtube_downloader.html", {"platform": "youtube"})

def api_downloader_info(request):
    """API endpoint to analyze URL and return media info."""
    if request.method != "POST":
        return JsonResponse({"success": False, "message": "Method not allowed"}, status=405)
    
    try:
        data = json.loads(request.body)
        url = data.get("url")
        if not url:
            return JsonResponse({"success": False, "message": "URL is required."}, status=400)
            
        if len(url) > 2000:
            return JsonResponse({"success": False, "message": "URL is too long."}, status=400)
            
        info = get_media_info(url)
        return JsonResponse({"success": True, **info})
    except ValueError as e:
        return JsonResponse({"success": False, "message": str(e)}, status=400)
    except Exception as e:
        logger.error(f"Downloader info error: {str(e)}")
        return JsonResponse({"success": False, "message": "Failed to analyze media."}, status=500)

import re

def sanitize_filename(name):
    clean = re.sub(r'[^a-zA-Z0-9\.\_\-]', '_', name)
    return clean[:200]

def api_downloader_download(request):
    """API endpoint to stream media download and clean up temp file."""
    if request.method != "POST":
        return JsonResponse({"success": False, "message": "Method not allowed"}, status=405)
        
    filepath = None
    try:
        data = json.loads(request.body)
        url = data.get("url")
        format_id = data.get("format_id")
        download_type = data.get("download_type", "video")
        audio_quality = data.get("audio_quality", "192")
        
        if not url:
            return JsonResponse({"success": False, "message": "URL is required."}, status=400)
            
        if len(url) > 2000:
            return JsonResponse({"success": False, "message": "URL is too long."}, status=400)
            
        filepath, filename = download_media_to_temp(url, format_id, download_type, audio_quality)
        safe_filename = sanitize_filename(filename) or f"downloaded_media.{'mp3' if download_type == 'audio' else 'mp4'}"
        
        # Stream the file and delete it after
        def file_iterator(path, chunk_size=8192):
            try:
                with open(path, 'rb') as f:
                    while True:
                        chunk = f.read(chunk_size)
                        if not chunk:
                            break
                        yield chunk
            finally:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except Exception as e:
                    logger.error(f"Failed to delete temp file {path}: {str(e)}")

        response = StreamingHttpResponse(file_iterator(filepath), content_type="application/octet-stream")
        response['Content-Disposition'] = f'attachment; filename="{safe_filename}"'
        response['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        return response
        
    except ValueError as e:
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
        return JsonResponse({"success": False, "message": str(e)}, status=400)
    except Exception as e:
        logger.error(f"Downloader download error: {str(e)}")
        if filepath and os.path.exists(filepath):
            try:
                os.remove(filepath)
            except:
                pass
        return JsonResponse({"success": False, "message": "Failed to download media."}, status=500)
