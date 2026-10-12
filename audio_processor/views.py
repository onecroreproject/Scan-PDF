import os
import mimetypes
import logging
from pathlib import Path
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST

from .utils import (
    process_audio,
    merge_audios,
    extract_audio_from_video,
    probe_media,
    SUPPORTED_AUDIO_EXTENSIONS,
    SUPPORTED_VIDEO_EXTENSIONS,
    SUPPORTED_OUTPUT_FORMATS,
)
from converter.utils import save_uploaded_file, format_download_name

logger = logging.getLogger(__name__)


class FileCleanupResponse(FileResponse):
    """FileResponse that removes the temp file from disk once stream is closed."""
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
        'description': 'Edit audio directly in your browser workflow with precise trimming, volume controls, and clean downloadable output.',
        'icon': 'music',
        'accept': '.mp3,.wav,.ogg,.m4a,.flac',
        'allowed_extensions': ['.mp3', '.wav', '.ogg', '.m4a', '.flac'],
        'category': 'audio-tools',
        'color': '#7c3aed',
        'gradient': 'from-purple-700 to-purple-500',
    },
    'merge-audio': {
        'title': 'Merge Audio',
        'description': 'Combine multiple audio files into one continuous track while keeping the order you choose.',
        'icon': 'combine',
        'accept': '.mp3,.wav,.ogg,.m4a,.flac',
        'allowed_extensions': ['.mp3', '.wav', '.ogg', '.m4a', '.flac'],
        'category': 'audio-tools',
        'multi_file': True,
        'color': '#7c3aed',
        'gradient': 'from-purple-700 to-purple-500',
    },
    'extract-audio-from-video': {
        'title': 'Extract Audio From Video',
        'description': 'Extract the audio track from your video and download it in your preferred supported audio format.',
        'icon': 'video',
        'accept': '.mp4,.mov,.avi,.mkv,.webm',
        'allowed_extensions': ['.mp4', '.mov', '.avi', '.mkv', '.webm'],
        'category': 'audio-tools',
        'color': '#7c3aed',
        'gradient': 'from-purple-700 to-purple-500',
    }
}


def editor_page(request):
    """Render the main audio editor interface."""
    from converter.seo_content import SEO_DATA
    tool = dict(AUDIO_TOOLS['audio-editor'])
    seo_extra = SEO_DATA.get('audio-editor', {})
    tool.update(seo_extra)
    context = {
        'tool': tool,
        'tool_slug': 'audio-editor',
        'seo_title': tool.get('seo_title', 'Audio Editor Online – Free MP3 WAV Audio Editing Tool | ScanPDF'),
        'seo_description': tool.get('seo_description', tool['description']),
        'seo_keywords': tool.get('seo_keywords', ''),
        'seo_h1': tool.get('seo_h1', 'Online Audio Editor'),
        'seo_intro': tool.get('seo_intro', tool['description']),
        's1': 'Online',
        'highlight': 'Audio Editor',
    }
    return render(request, 'audio_processor/editor.html', context)


def merge_page(request):
    """Render dedicated merge-audio page."""
    from converter.seo_content import SEO_DATA
    tool = dict(AUDIO_TOOLS['merge-audio'])
    seo_extra = SEO_DATA.get('merge-audio', {})
    tool.update(seo_extra)
    context = {
        'tool': tool,
        'tool_slug': 'merge-audio',
        'seo_title': tool.get('seo_title', 'Merge Audio Files Online – Free Audio Merger | ScanPDF'),
        'seo_description': tool.get('seo_description', tool['description']),
        'seo_keywords': tool.get('seo_keywords', ''),
        'seo_h1': tool.get('seo_h1', 'Merge Audio Files Online'),
        'seo_intro': tool.get('seo_intro', tool['description']),
        's1': 'Merge',
        'highlight': 'Audio Files Online',
    }
    return render(request, 'audio_processor/merge_audio.html', context)


def extract_page(request):
    """Render dedicated video-to-audio extraction page."""
    from converter.seo_content import SEO_DATA
    tool = dict(AUDIO_TOOLS['extract-audio-from-video'])
    seo_extra = SEO_DATA.get('extract-audio-from-video', {})
    tool.update(seo_extra)
    context = {
        'tool': tool,
        'tool_slug': 'extract-audio-from-video',
        'seo_title': tool.get('seo_title', 'Extract Audio from Video Online – Free Video to Audio Tool | ScanPDF'),
        'seo_description': tool.get('seo_description', tool['description']),
        'seo_keywords': tool.get('seo_keywords', ''),
        'seo_h1': tool.get('seo_h1', 'Extract Audio from Video'),
        'seo_intro': tool.get('seo_intro', tool['description']),
        's1': 'Extract Audio from',
        'highlight': 'Video Online',
    }
    return render(request, 'audio_processor/extract_audio.html', context)


@csrf_exempt
@require_POST
def process_audio_api(request):
    """Universal API endpoint supporting audio-editor, merge-audio, and extract-audio-from-video."""
    mode = request.POST.get('mode', 'edit').strip().lower()
    output_format = request.POST.get('format', 'mp3').lower().strip()
    if output_format not in SUPPORTED_OUTPUT_FORMATS:
        output_format = 'mp3'

    # ── 1. MERGE MODE ──────────────────────────────────────────────
    if mode == 'merge':
        files = request.FILES.getlist('files')
        if not files:
            files = [request.FILES.get('file')] if 'file' in request.FILES else []

        if len(files) < 2:
            return JsonResponse({'error': 'Please upload at least 2 audio files to merge.'}, status=400)

        # Validate file extensions
        for f in files:
            ext = os.path.splitext(f.name)[1].lower()
            if ext not in SUPPORTED_AUDIO_EXTENSIONS:
                allowed_str = ", ".join(sorted(SUPPORTED_AUDIO_EXTENSIONS))
                return JsonResponse({'error': f'Invalid file "{f.name}". Supported audio formats: {allowed_str}'}, status=400)

        input_paths = []
        try:
            for f in files:
                input_paths.append(save_uploaded_file(f))

            output_path = merge_audios(input_paths, files[0].name, target_format=output_format)
            ext = output_format
            return create_cleanup_response(output_path, filename=f"merged-audio.{ext}")
        except ValueError as e:
            return JsonResponse({'error': str(e)}, status=400)
        except Exception as e:
            logger.error(f"Merge audio error: {e}", exc_info=True)
            return JsonResponse({'error': f"Merge audio failed: {str(e)}"}, status=500)
        finally:
            for p in input_paths:
                if os.path.exists(p):
                    try: os.remove(p)
                    except OSError: pass

    # ── 2. EXTRACT MODE ────────────────────────────────────────────
    if mode == 'extract':
        video_file = request.FILES.get('video_file') or request.FILES.get('file')
        if not video_file:
            return JsonResponse({'error': 'Please upload a video file.'}, status=400)

        ext = os.path.splitext(video_file.name)[1].lower()
        if ext not in SUPPORTED_VIDEO_EXTENSIONS:
            allowed_str = ", ".join(sorted(SUPPORTED_VIDEO_EXTENSIONS))
            return JsonResponse({'error': f'Invalid video format "{ext}". Supported video formats: {allowed_str}'}, status=400)

        extract_mode = request.POST.get('extract_mode', 'full').strip().lower()
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
            base = os.path.splitext(video_file.name)[0]
            return create_cleanup_response(output_path, filename=f"{base}_audio.{output_format}")
        except ValueError as e:
            return JsonResponse({'error': str(e)}, status=400)
        except Exception as e:
            logger.error(f"Extract audio error: {e}", exc_info=True)
            return JsonResponse({'error': f"Audio extraction failed: {str(e)}"}, status=500)
        finally:
            if os.path.exists(input_path):
                try: os.remove(input_path)
                except OSError: pass

    # ── 3. EDIT MODE ───────────────────────────────────────────────
    uploaded_file = request.FILES.get('file') or request.FILES.get('audio_file')
    if not uploaded_file:
        return JsonResponse({'error': 'Please upload an audio file to edit.'}, status=400)

    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext not in SUPPORTED_AUDIO_EXTENSIONS:
        allowed_str = ", ".join(sorted(SUPPORTED_AUDIO_EXTENSIONS))
        return JsonResponse({'error': f'Invalid audio format "{ext}". Supported audio formats: {allowed_str}'}, status=400)

    input_path = save_uploaded_file(uploaded_file)

    tool_params = {
        'start': request.POST.get('start', 0),
        'end': request.POST.get('end', ''),
        'volume': request.POST.get('volume', 100),
        'fade_in': request.POST.get('fade_in', 0),
        'fade_out': request.POST.get('fade_out', 0),
        'format': output_format,
        'bitrate': request.POST.get('bitrate', '192k'),
    }

    try:
        output_path = process_audio(input_path, uploaded_file.name, tool_params)
        base = os.path.splitext(uploaded_file.name)[0]
        return create_cleanup_response(output_path, filename=f"{base}_edited.{output_format}")
    except ValueError as e:
        return JsonResponse({'error': str(e)}, status=400)
    except Exception as e:
        logger.error(f"Audio editor error: {e}", exc_info=True)
        return JsonResponse({'error': f"Audio editing failed: {str(e)}"}, status=500)
    finally:
        if os.path.exists(input_path):
            try: os.remove(input_path)
            except OSError: pass
