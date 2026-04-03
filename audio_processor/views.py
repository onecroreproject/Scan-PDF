import os
import mimetypes
from pathlib import Path
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from django import forms
from .utils import process_audio, merge_audios, format_download_name

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

AUDIO_TOOLS = {
    'trim-audio': {
        'title': 'Trim Audio',
        'description': 'Cut specific parts of your audio files with fade in/out effects.',
        'icon': 'scissors',
        'accept': '.mp3,.wav,.m4a,.flac',
        'allowed_extensions': ['.mp3', '.wav', '.m4a', '.flac'],
        'category': 'audio',
        'color': '#3b82f6',
        'gradient': 'from-blue-500 to-indigo-600',
    },
    'change-volume': {
        'title': 'Change Volume',
        'description': 'Easily increase or decrease the loudness of your audio files.',
        'icon': 'volume-2',
        'accept': '.mp3,.wav',
        'allowed_extensions': ['.mp3', '.wav'],
        'category': 'audio',
        'color': '#ef4444',
        'gradient': 'from-red-500 to-orange-600',
    },
    'change-speed': {
        'title': 'Change Speed',
        'description': 'Speed up or slow down your audio tracks effortlessly.',
        'icon': 'zap',
        'accept': '.mp3,.wav',
        'allowed_extensions': ['.mp3', '.wav'],
        'category': 'audio',
        'color': '#f59e0b',
        'gradient': 'from-amber-400 to-orange-500',
    },
    'reverse-audio': {
        'title': 'Reverse Audio',
        'description': 'Instantly flip your audio files to play backwards.',
        'icon': 'rotate-ccw',
        'accept': '.mp3,.wav',
        'allowed_extensions': ['.mp3', '.wav'],
        'category': 'audio',
        'color': '#6366f1',
        'gradient': 'from-indigo-400 to-violet-600',
    },
    'merge-audio': {
        'title': 'Merge Audio',
        'description': 'Combine multiple audio tracks into a single seamless file.',
        'icon': 'combine',
        'accept': '.mp3,.wav',
        'allowed_extensions': ['.mp3', '.wav'],
        'category': 'audio',
        'color': '#10b981',
        'gradient': 'from-emerald-400 to-teal-600',
        'multi_file': True,
    },
    'video-to-audio': {
        'title': 'Video to MP3',
        'description': 'Extract high-quality audio from your video files instantly.',
        'icon': 'file-audio',
        'accept': '.mp4,.mov,.avi,.mkv',
        'allowed_extensions': ['.mp4', '.mov', '.avi', '.mkv'],
        'category': 'audio',
        'color': '#f43f5e',
        'gradient': 'from-rose-500 to-pink-600',
    }
}

def tool_page(request, tool_slug):
    if tool_slug not in AUDIO_TOOLS: raise Http404("Audio tool not found")
    tool = AUDIO_TOOLS[tool_slug]
    return render(request, 'audio_processor/tool_detail.html', {
        'tool': tool,
        'tool_slug': tool_slug,
        'page_title': f'{tool["title"]} — ScanPDF',
    })

@csrf_exempt
@require_POST
def process_tool(request, tool_slug):
    if tool_slug not in AUDIO_TOOLS:
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

        output_path = None
        
        # Build params
        params = request.POST.dict()
        params['tool'] = tool_slug

        if tool_slug == 'merge-audio':
            output_path = merge_audios(input_paths, original_name)
        else:
            output_path = process_audio(input_paths[0], original_name, params)
        
        if output_path and os.path.exists(output_path):
            return create_cleanup_response(output_path)
        else:
            return JsonResponse({'error': 'Failed to process audio.'}, status=500)

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    finally:
        for p in input_paths:
            if os.path.exists(p):
                try: os.remove(p)
                except: pass
