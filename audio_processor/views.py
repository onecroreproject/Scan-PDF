import os
import mimetypes
from pathlib import Path
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from .utils import process_audio
from converter.utils import save_uploaded_file, format_download_name

class FileCleanupResponse(FileResponse):
    def __init__(self, file_path, *args, **kwargs):
        self._temp_file_path = file_path
        file_handle = open(file_path, 'rb')
        super().__init__(file_handle, *args, **kwargs)

    def close(self):
        super().close()
        if self._temp_file_path and os.path.exists(self._temp_file_path):
            try:
                os.remove(self._temp_file_path)
            except OSError:
                pass

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
    'audio-editor': {
        'title': 'Audio Editor',
        'description': 'Professional audio editor to trim, cut, change volume, speed, pitch and apply equalizer effects.',
        'icon': 'music',
        'accept': '.mp3,.wav,.ogg,.m4a,.flac',
        'allowed_extensions': ['.mp3', '.wav', '.ogg', '.m4a', '.flac'],
        'category': 'audio-tools',
        'color': '#10b981',
        'gradient': 'from-emerald-500 to-teal-600',
    }
}

def editor_page(request):
    """Render the main audio editor interface."""
    tool = AUDIO_TOOLS['audio-editor']
    context = {
        'tool': tool,
        'tool_slug': 'audio-editor',
        'page_title': 'Audio Editor — ScanPDF',
    }
    return render(request, 'audio_processor/editor.html', context)

@csrf_exempt
@require_POST
def process_audio_api(request):
    """API endpoint to process audio files."""
    uploaded_file = request.FILES.get('file')
    if not uploaded_file:
        return JsonResponse({'error': 'No file uploaded'}, status=400)
    
    input_path = save_uploaded_file(uploaded_file)
    
    # Get tool parameters from POST
    tool_params = {
        'tool': request.POST.get('tool', 'trim-audio'),
        'start': request.POST.get('start', 0),
        'end': request.POST.get('end', 0),
        'fade_in': request.POST.get('fade_in', 0),
        'fade_out': request.POST.get('fade_out', 0),
        'volume': request.POST.get('volume', 100),
        'speed': request.POST.get('speed', 1.0),
        'pitch': request.POST.get('pitch', 0), # Pitch might need separate handling
        'preset': request.POST.get('preset', 'none'),
        'format': request.POST.get('format', 'mp3'),
    }
    
    # Optional: handle individual equalizer bands if preset is not used
    # The current utils.py doesn't handle bands yet, I should update it.

    try:
        output_path = process_audio(input_path, uploaded_file.name, tool_params)
        return create_cleanup_response(output_path, filename=uploaded_file.name)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    finally:
        if os.path.exists(input_path):
            try: os.remove(input_path)
            except: pass
