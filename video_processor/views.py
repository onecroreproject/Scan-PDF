import os
import mimetypes
from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from .utils import save_upload, convert_video_moviepy, cleanup_old_files

# ═══════════════════════════════════════════════════════════════
# FILE CLEANUP RESPONSE
# ═══════════════════════════════════════════════════════════════

class FileCleanupResponse(FileResponse):
    """FileResponse that deletes the file after it's closed."""
    def __init__(self, file_path, *args, **kwargs):
        self._temp_file_path = file_path
        super().__init__(open(file_path, 'rb'), *args, **kwargs)

    def close(self):
        super().close()
        if os.path.exists(self._temp_file_path):
            try:
                os.remove(self._temp_file_path)
            except OSError:
                pass

def converter_page(request):
    """Render the main Video Converter page."""
    # Periodic cleanup of old files (every time someone visits the page)
    cleanup_old_files()
    return render(request, 'video_processor/converter.html')

@csrf_exempt
def convert_video_api(request):
    """API endpoint to process video conversion."""
    if request.method != 'POST':
        return JsonResponse({'error': 'POST request required'}, status=405)
    
    video_file = request.FILES.get('video')
    output_format = request.POST.get('format')
    
    if not video_file:
        return JsonResponse({'error': 'No video file uploaded'}, status=400)
    
    if not output_format:
        return JsonResponse({'error': 'Output format not selected'}, status=400)
    
    # 1. Validate File
    allowed_inputs = ['.mp4', '.mov', '.avi', '.mkv', '.webm']
    ext = os.path.splitext(video_file.name)[1].lower()
    if ext not in allowed_inputs:
        return JsonResponse({'error': f'Unsupported input format: {ext}'}, status=400)
    
    if video_file.size > 50 * 1024 * 1024: # 50MB limit
        return JsonResponse({'error': 'File too large (Max 50MB)'}, status=400)

    # 2. Save Upload
    input_path = None
    try:
        input_path = save_upload(video_file)
        
        # 3. Convert using MoviePy
        output_path = convert_video_moviepy(input_path, output_format)
        
        # 4. Return file as response (which will auto-delete on close)
        content_type, _ = mimetypes.guess_type(output_path)
        response = FileCleanupResponse(output_path, content_type=content_type)
        response['Content-Disposition'] = f'attachment; filename="converted_{video_file.name.split(".")[0]}.{output_format}"'
        return response
        
    except Exception as e:
        return JsonResponse({'error': f'Conversion failed: {str(e)}'}, status=500)
    finally:
        # 5. Clean up original upload
        if input_path and os.path.exists(input_path):
            try:
                os.remove(input_path)
            except:
                pass
