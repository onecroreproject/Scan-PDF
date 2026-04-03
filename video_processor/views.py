import os
import mimetypes
from pathlib import Path
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from django import forms
from .utils import process_video, merge_videos, format_download_name

class FileCleanupResponse(FileResponse):
    def __init__(self, file_path, *args, **kwargs):
        self._temp_file_path = file_path
        file_handle = open(file_path, 'rb')
        super().__init__(file_handle, *args, **kwargs)

    def close(self):
        super().close()
        if self._temp_file_path and os.path.exists(self._temp_file_path):
            try: os.remove(self._temp_file_path)
            except OSError: pass

def create_cleanup_response(file_path, content_type=None, filename=None):
    if not content_type:
        content_type, _ = mimetypes.guess_type(file_path)
        content_type = content_type or 'application/octet-stream'
    
    raw_name = filename or os.path.basename(file_path)
    final_filename = format_download_name(raw_name)

    response = FileCleanupResponse(file_path, content_type=content_type)
    response['Content-Disposition'] = f'attachment; filename="{final_filename}"'
    return response

VIDEO_TOOLS = {
    'trim-video': {
        'title': 'Trim Video',
        'description': 'Cut specific segments of your video files with precision.',
        'icon': 'scissors',
        'accept': '.mp4,.mov,.avi,.mkv',
        'allowed_extensions': ['.mp4', '.mov', '.avi', '.mkv'],
        'category': 'video',
        'color': '#8b5cf6',
        'gradient': 'from-violet-500 to-purple-700',
    },
    'crop-video': {
        'title': 'Crop Video',
        'description': 'Re-adjust the framing of your video clips.',
        'icon': 'crop',
        'accept': '.mp4,.mov',
        'allowed_extensions': ['.mp4', '.mov'],
        'category': 'video',
        'color': '#10b981',
        'gradient': 'from-emerald-500 to-teal-700',
    },
    'rotate-video': {
        'title': 'Rotate Video',
        'description': 'Flip or rotate your videos in any direction.',
        'icon': 'rotate-cw',
        'accept': '.mp4,.mov',
        'allowed_extensions': ['.mp4', '.mov'],
        'category': 'video',
        'color': '#f43f5e',
        'gradient': 'from-rose-500 to-pink-700',
    },
    'change-speed-video': {
        'title': 'Change Speed',
        'description': 'Accelerate or slow down your video clips.',
        'icon': 'zap',
        'accept': '.mp4,.mov',
        'allowed_extensions': ['.mp4', '.mov'],
        'category': 'video',
        'color': '#f59e0b',
        'gradient': 'from-amber-400 to-orange-600',
    },
    'merge-video': {
        'title': 'Merge Video',
        'description': 'Combine several video clips into one masterpiece.',
        'icon': 'combine',
        'accept': '.mp4,.mov',
        'allowed_extensions': ['.mp4', '.mov'],
        'category': 'video',
        'color': '#3b82f6',
        'gradient': 'from-blue-500 to-indigo-700',
        'multi_file': True,
    }
}

def tool_page(request, tool_slug):
    if tool_slug not in VIDEO_TOOLS: raise Http404("Video tool not found")
    tool = VIDEO_TOOLS[tool_slug]
    return render(request, 'video_processor/tool_detail.html', {
        'tool': tool,
        'tool_slug': tool_slug,
        'page_title': f'{tool["title"]} — ScanPDF',
    })

@csrf_exempt
@require_POST
def process_tool(request, tool_slug):
    if tool_slug not in VIDEO_TOOLS:
        return JsonResponse({'error': 'Invalid tool.'}, status=400)

    files = request.FILES.getlist('files')
    if not files:
        file = request.FILES.get('file')
        files = [file] if file else []

    if not files:
        return JsonResponse({'error': 'No file uploaded.'}, status=400)

    input_paths = []
    try:
        from converter.utils import save_uploaded_file
        original_name = files[0].name
        for f in files:
            p = save_uploaded_file(f)
            input_paths.append(p)

        params = request.POST.dict()
        params['tool'] = tool_slug

        if tool_slug == 'merge-video':
            output_path = merge_videos(input_paths, original_name)
        else:
            output_path = process_video(input_paths[0], original_name, params)
        
        if output_path and os.path.exists(output_path):
            return create_cleanup_response(output_path)
        else:
            return JsonResponse({'error': 'Failed to process video.'}, status=500)

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    finally:
        for p in input_paths:
            if os.path.exists(p):
                try: os.remove(p)
                except: pass
