import os
import mimetypes
from pathlib import Path
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from .utils import process_audio, merge_audios, extract_audio_from_video
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

def merge_page(request):
    """Render dedicated merge-audio page."""
    context = {
        'tool_slug': 'audio-merge',
        'tool': {
            'title': 'Merge Audio',
            'description': 'Combine multiple audio files into one output with clear file review and remove options.',
        },
        'page_title': 'Merge Audio — ScanPDF',
    }
    return render(request, 'audio_processor/merge_audio.html', context)

def extract_page(request):
    """Render dedicated video-to-audio extraction page."""
    context = {
        'tool_slug': 'extract-audio',
        'tool': {
            'title': 'Extract Audio From Video',
            'description': 'Upload a video and extract either full audio or a custom start-to-end range.',
        },
        'page_title': 'Extract Audio — ScanPDF',
    }
    return render(request, 'audio_processor/extract_audio.html', context)

@csrf_exempt
@require_POST
def process_audio_api(request):
    """API endpoint to process audio files in editor, merge, or extract modes."""
    mode = request.POST.get('mode', 'edit')
    allowed_output_formats = {'mp3', 'wav', 'ogg', 'm4a', 'flac'}
    output_format = request.POST.get('format', 'mp3').lower()
    if output_format not in allowed_output_formats:
        return JsonResponse({'error': 'Unsupported output format.'}, status=400)

    if mode == 'merge':
        files = request.FILES.getlist('files')
        if len(files) < 2:
            return JsonResponse({'error': 'Please upload at least 2 audio files to merge.'}, status=400)
        input_paths = [save_uploaded_file(f) for f in files]
        try:
            output_path = merge_audios(input_paths, files[0].name, target_format=output_format)
            merged_name = f"{Path(files[0].name).stem}_merged.{output_format}"
            return create_cleanup_response(output_path, filename=merged_name)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)
        finally:
            for path in input_paths:
                if os.path.exists(path):
                    try:
                        os.remove(path)
                    except OSError:
                        pass

    if mode == 'extract':
        video_file = request.FILES.get('video_file')
        if not video_file:
            return JsonResponse({'error': 'Please upload a video file.'}, status=400)
        extract_mode = request.POST.get('extract_mode', 'full')
        start = request.POST.get('start', 0)
        end = request.POST.get('end', '')
        input_path = save_uploaded_file(video_file)
        try:
            output_path = extract_audio_from_video(
                input_path,
                video_file.name,
                target_format=output_format,
                extract_mode=extract_mode,
                start=start,
                end=end,
            )
            extracted_name = f"{Path(video_file.name).stem}_audio.{output_format}"
            return create_cleanup_response(output_path, filename=extracted_name)
        except ValueError as e:
            return JsonResponse({'error': str(e)}, status=400)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)
        finally:
            if os.path.exists(input_path):
                try:
                    os.remove(input_path)
                except OSError:
                    pass

    uploaded_file = request.FILES.get('file')
    if not uploaded_file:
        return JsonResponse({'error': 'No file uploaded'}, status=400)
    input_path = save_uploaded_file(uploaded_file)

    preset = request.POST.get('preset', 'none').strip().lower()
    preset_aliases = {'full-bass': 'bass-boost', 'full-treble': 'treble-boost'}
    preset = preset_aliases.get(preset, preset)
    allowed_presets = {'none', 'classic', 'dance', 'club', 'bass-boost', 'treble-boost', 'pop', 'rock'}
    if preset not in allowed_presets:
        return JsonResponse({'error': 'Unsupported equalizer preset.'}, status=400)

    # Ringtone quick preset
    is_ringtone = request.POST.get('ringtone_mode', 'false') == 'true'
    start = request.POST.get('start', 0)
    end = request.POST.get('end', 0)
    if is_ringtone:
        try:
            start_f = max(0.0, float(start or 0))
        except (TypeError, ValueError):
            start_f = 0.0
        end_f = start_f + 40.0
        start = str(start_f)
        end = str(end_f)

    tool_params = {
        'start': start,
        'end': end,
        'fade_in': request.POST.get('fade_in', 0),
        'fade_out': request.POST.get('fade_out', 0),
        'volume': request.POST.get('volume', 100),
        'speed': request.POST.get('speed', 1.0),
        'pitch': request.POST.get('pitch', 0),
        'preset': preset,
        'format': output_format,
        'reverse': request.POST.get('reverse', 'false'),
        'bitrate': request.POST.get('bitrate', '192k'),
    }

    try:
        output_path = process_audio(input_path, uploaded_file.name, tool_params)
        download_name = f"{Path(uploaded_file.name).stem}_edited.{output_format}"
        return create_cleanup_response(output_path, filename=download_name)
    except ValueError as e:
        return JsonResponse({'error': str(e)}, status=400)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    finally:
        if os.path.exists(input_path):
            try:
                os.remove(input_path)
            except OSError:
                pass
