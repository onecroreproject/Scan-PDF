import json
import os
import logging
from django.shortcuts import render
from django.http import JsonResponse, StreamingHttpResponse, Http404, HttpResponse
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
    """API endpoint to prepare download (POST) or stream file (GET)."""
    if request.method == "GET":
        file_id = request.GET.get("file_id")
        filename = request.GET.get("filename", "download")
        
        if not file_id or not re.match(r'^dl_[a-f0-9]+\.[a-zA-Z0-9]+$', file_id):
            return HttpResponse("Invalid file ID", status=400)
            
        import tempfile
        filepath = os.path.join(tempfile.gettempdir(), file_id)
        if not os.path.exists(filepath):
            return HttpResponse("File expired or not found", status=404)
            
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

        file_size = os.path.getsize(filepath)
        content_type = "audio/mpeg" if filename.lower().endswith(".mp3") else "application/octet-stream"
        
        logger.info(f"YOUTUBE MP3 DIAGNOSTICS: Starting response. Content-Type: {content_type}, Size: {file_size}")

        response = StreamingHttpResponse(file_iterator(filepath), content_type=content_type)
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        response['Content-Length'] = str(file_size)
        response['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        return response

    if request.method != "POST":
        return JsonResponse({"success": False, "message": "Method not allowed"}, status=405)
        
    filepath = None
    try:
        data = json.loads(request.body)
        url = data.get("url")
        format_id = data.get("format_id")
        download_type = data.get("download_type", "video")
        audio_quality = data.get("audio_quality", "192")
        media_url = data.get("media_url")
        image_format = data.get("image_format")
        
        if not url:
            return JsonResponse({"success": False, "message": "URL is required."}, status=400)
            
        if len(url) > 2000:
            return JsonResponse({"success": False, "message": "URL is too long."}, status=400)
            
        if download_type == "image" and media_url:
            from urllib.parse import urlparse
            hostname = urlparse(media_url).hostname or ""
            if not (hostname.endswith('.cdninstagram.com') or hostname.endswith('.fbcdn.net') or hostname.endswith('.threads.net') or hostname.endswith('.twimg.com')):
                return JsonResponse({"success": False, "message": "Invalid media URL host."}, status=400)
                
            import urllib.request
            req = urllib.request.Request(media_url, headers={'User-Agent': 'Mozilla/5.0'})
            try:
                cdn_resp = urllib.request.urlopen(req, timeout=15)
            except Exception as e:
                return JsonResponse({"success": False, "message": "Failed to fetch image."}, status=400)
                
            content_type = cdn_resp.headers.get('Content-Type', '')
            if not content_type.startswith('image/'):
                return JsonResponse({"success": False, "message": "Invalid content type."}, status=400)

            import io
            image_data = cdn_resp.read()
            cdn_resp.close()

            if image_format in ["jpg", "png"]:
                from PIL import Image
                try:
                    img = Image.open(io.BytesIO(image_data))
                    out_io = io.BytesIO()
                    if image_format == "jpg":
                        if img.mode in ("RGBA", "P", "LA"):
                            bg = Image.new("RGB", img.size, (255, 255, 255))
                            if img.mode in ("RGBA", "LA"):
                                bg.paste(img, mask=img.split()[-1])
                            else:
                                bg.paste(img)
                            img = bg
                        elif img.mode != "RGB":
                            img = img.convert("RGB")
                        img.save(out_io, format="JPEG", quality=95)
                        content_type = "image/jpeg"
                        ext = "jpg"
                    elif image_format == "png":
                        img.save(out_io, format="PNG")
                        content_type = "image/png"
                        ext = "png"
                    image_data = out_io.getvalue()
                except Exception as e:
                    logger.error(f"Image conversion failed: {str(e)}")
                    return JsonResponse({"success": False, "message": "Failed to convert image."}, status=500)
            else:
                ext = "jpg"
                if "png" in content_type: ext = "png"
                elif "webp" in content_type: ext = "webp"

            import tempfile, uuid
            temp_dir = tempfile.gettempdir()
            file_id = f"dl_{uuid.uuid4().hex}.{ext}"
            filepath = os.path.join(temp_dir, file_id)
            with open(filepath, 'wb') as f:
                f.write(image_data)
                
            prefix = "media"
            if 'twimg.com' in hostname: prefix = "twitter"
            elif 'cdninstagram.com' in hostname: prefix = "instagram"
            elif 'fbcdn.net' in hostname: prefix = "facebook"
            elif 'threads.net' in hostname: prefix = "threads"
                
            return JsonResponse({
                "success": True, 
                "file_id": file_id,
                "filename": f"{prefix}-image.{ext}"
            })
            
        filepath, filename = download_media_to_temp(url, format_id, download_type, audio_quality)
        safe_filename = sanitize_filename(filename) or f"downloaded_media.{'mp3' if download_type == 'audio' else 'mp4'}"
        
        file_id = os.path.basename(filepath)
        return JsonResponse({
            "success": True, 
            "file_id": file_id,
            "filename": safe_filename
        })
        
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
