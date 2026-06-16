"""
Video Processor Views — All 10 Tools
Direct FFmpeg execution with security validation and temp cleanup.
"""
import os
import json
import uuid
import mimetypes
import logging
import threading
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from django.shortcuts import render
from django.http import JsonResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from django.core.exceptions import ValidationError
from django.conf import settings

from .utils import (
    save_upload, safe_remove, cleanup_old_files, get_uploads_dir,
    save_chunk, assemble_chunks, get_chunk_status, FileCleanupResponse
)
from .validators import (
    validate_video, validate_image, validate_audio, validate_subtitle,
    sanitize_filename
)
from .ffmpeg_helpers import (
    run_ffmpeg,
    build_convert_command,
    build_image_to_video_command,
    build_compress_command,
    build_merge_command,
    build_gif_command,
    build_audio_extract_command,
    build_watermark_command,
    build_subtitle_command,
    build_text_overlay_command,
    build_trim_command,
    build_rotate_command,
    build_resize_command,
    build_crop_command,
    build_speed_command,
    build_mute_command,
    build_replace_audio_command,
    build_cut_command,
    get_video_info,
)

logger = logging.getLogger(__name__)
MAX_SYNC_MB = 512  # Files above this should ideally use Celery; kept for direct processing

# Background job tracking (for sync compression fallback)
_background_jobs = {}

def _run_compression_background(input_path, output_path, options, job_id, use_chunk_path):
    """Run FFmpeg compression in background thread."""
    try:
        logger.info(f"[{job_id}] Starting background compression")
        cmd = build_compress_command(input_path, output_path, options)
        run_ffmpeg(cmd, timeout=1800)
        logger.info(f"[{job_id}] Compression complete: {output_path}")
        _background_jobs[job_id] = {'status': 'complete', 'output_path': output_path}
    except Exception as e:
        logger.exception(f"[{job_id}] Background compression failed")
        safe_remove(output_path)
        _background_jobs[job_id] = {'status': 'failed', 'error': str(e)}
    finally:
        # Only delete input if it was a direct upload
        if not use_chunk_path:
            safe_remove(input_path)

# ═══════════════════════════════════════════════════════════════
# SHARED HELPERS
# ═══════════════════════════════════════════════════════════════

def _cleanup_response(output_path, filename=None):
    """Build a FileCleanupResponse with proper headers."""
    content_type, _ = mimetypes.guess_type(output_path)
    if not content_type:
        content_type = 'application/octet-stream'
    resp = FileCleanupResponse(output_path, content_type=content_type, filename=filename)
    return resp.build_response()


def _get_output_path(ext):
    """Generate a unique output path in the temp outputs directory."""
    from .utils import get_outputs_dir
    return os.path.join(get_outputs_dir(), f"{uuid.uuid4().hex}.{ext.lstrip('.')}")


def _handle_uploaded_video(request, field_name='video', max_mb=2048):
    """Validate, save, and return path for a video upload."""
    file_obj = request.FILES.get(field_name)
    if not file_obj:
        raise ValidationError("No file uploaded")
    validate_video(file_obj, max_mb)
    return save_upload(file_obj)


def _handle_uploaded_image(request, field_name='image', max_mb=100):
    file_obj = request.FILES.get(field_name)
    if not file_obj:
        raise ValidationError("No image uploaded")
    validate_image(file_obj, max_mb)
    return save_upload(file_obj)


def _handle_uploaded_audio(request, field_name='audio', max_mb=500):
    file_obj = request.FILES.get(field_name)
    if not file_obj:
        raise ValidationError("No audio uploaded")
    validate_audio(file_obj, max_mb)
    return save_upload(file_obj)


def _handle_uploaded_subtitle(request, field_name='subtitle', max_mb=10):
    file_obj = request.FILES.get(field_name)
    if not file_obj:
        raise ValidationError("No subtitle file uploaded")
    validate_subtitle(file_obj, max_mb)
    return save_upload(file_obj)


def _json_error(message, status=400):
    return JsonResponse({'error': message}, status=status)


def apply_image_edits(image_path, edits):
    """
    Apply image edits using Pillow (rotation, flip, filters, etc.).
    Returns path to edited image (original is preserved).
    """
    if not edits or not isinstance(edits, dict):
        return image_path
    
    # Check if any edits are actually applied
    has_edits = any([
        edits.get('rotation', 0) != 0,
        edits.get('flipH', False),
        edits.get('flipV', False),
        edits.get('zoom', 100) != 100,
        edits.get('brightness', 0) != 0,
        edits.get('contrast', 0) != 0,
        edits.get('saturation', 0) != 0,
        edits.get('blur', 0) != 0,
        edits.get('opacity', 100) != 100,
        edits.get('filter', 'none') != 'none'
    ])
    
    if not has_edits:
        return image_path
    
    try:
        from .utils import get_outputs_dir
        img = Image.open(image_path)
        
        # Convert to RGBA if needed for opacity
        if edits.get('opacity', 100) != 100:
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
        
        # Rotation
        rotation = edits.get('rotation', 0)
        if rotation:
            img = img.rotate(rotation, expand=True)
        
        # Flip
        if edits.get('flipH', False):
            img = img.transpose(Image.FLIP_LEFT_RIGHT)
        if edits.get('flipV', False):
            img = img.transpose(Image.FLIP_TOP_BOTTOM)
        
        # Zoom (resize)
        zoom = edits.get('zoom', 100) / 100
        if zoom != 1.0:
            new_size = (int(img.width * zoom), int(img.height * zoom))
            img = img.resize(new_size, Image.Resampling.LANCZOS)
        
        # Brightness
        brightness = edits.get('brightness', 0)
        if brightness != 0:
            enhancer = ImageEnhance.Brightness(img)
            factor = 1.0 + (brightness / 100.0)
            img = enhancer.enhance(max(0.1, factor))
        
        # Contrast
        contrast = edits.get('contrast', 0)
        if contrast != 0:
            enhancer = ImageEnhance.Contrast(img)
            factor = 1.0 + (contrast / 100.0)
            img = enhancer.enhance(max(0.1, factor))
        
        # Saturation
        saturation = edits.get('saturation', 0)
        if saturation != 0:
            enhancer = ImageEnhance.Color(img)
            factor = 1.0 + (saturation / 100.0)
            img = enhancer.enhance(max(0.1, factor))
        
        # Blur
        blur = edits.get('blur', 0)
        if blur > 0:
            img = img.filter(ImageFilter.GaussianBlur(radius=blur))
        
        # Opacity
        opacity = edits.get('opacity', 100)
        if opacity != 100 and img.mode == 'RGBA':
            alpha = img.split()[3]
            alpha = alpha.point(lambda p: p * opacity / 100)
            img.putalpha(alpha)
        
        # Filter presets - apply after all other edits
        filter_type = edits.get('filter', 'none')
        if filter_type != 'none':
            try:
                if filter_type == 'grayscale':
                    if img.mode == 'RGBA':
                        r, g, b, a = img.split()
                        gray_img = ImageOps.grayscale(Image.merge('RGB', (r, g, b)))
                        img = Image.new('RGBA', img.size)
                        img.paste(gray_img.convert('RGBA'), (0, 0), a)
                    else:
                        img = ImageOps.grayscale(img)
                elif filter_type == 'sepia':
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    pixels = img.load()
                    for y in range(img.height):
                        for x in range(img.width):
                            r, g, b = pixels[x, y][:3]
                            tr = int(0.393 * r + 0.769 * g + 0.189 * b)
                            tg = int(0.349 * r + 0.686 * g + 0.168 * b)
                            tb = int(0.272 * r + 0.534 * g + 0.131 * b)
                            pixels[x, y] = (min(255, tr), min(255, tg), min(255, tb))
                elif filter_type == 'invert':
                    if img.mode == 'RGBA':
                        r, g, b, a = img.split()
                        rgb = Image.merge('RGB', (r, g, b))
                        rgb = ImageOps.invert(rgb)
                        img = Image.merge('RGBA', (*rgb.split(), a))
                    else:
                        img = ImageOps.invert(img)
                elif filter_type == 'warm':
                    if img.mode not in ('RGB', 'RGBA'):
                        img = img.convert('RGB')
                    enhancer = ImageEnhance.Color(img)
                    img = enhancer.enhance(1.2)
                    # Warm tone via slight red increase and blue decrease
                    if img.mode == 'RGB':
                        r, g, b = img.split()
                        r = r.point(lambda x: min(255, int(x * 1.1)))
                        b = b.point(lambda x: max(0, int(x * 0.9)))
                        img = Image.merge('RGB', (r, g, b))
                    elif img.mode == 'RGBA':
                        r, g, b, a = img.split()
                        r = r.point(lambda x: min(255, int(x * 1.1)))
                        b = b.point(lambda x: max(0, int(x * 0.9)))
                        img = Image.merge('RGBA', (r, g, b, a))
                elif filter_type == 'cool':
                    if img.mode not in ('RGB', 'RGBA'):
                        img = img.convert('RGB')
                    enhancer = ImageEnhance.Color(img)
                    img = enhancer.enhance(0.9)
                    # Cool tone via blue increase and red decrease
                    if img.mode == 'RGB':
                        r, g, b = img.split()
                        r = r.point(lambda x: max(0, int(x * 0.9)))
                        b = b.point(lambda x: min(255, int(x * 1.1)))
                        img = Image.merge('RGB', (r, g, b))
                    elif img.mode == 'RGBA':
                        r, g, b, a = img.split()
                        r = r.point(lambda x: max(0, int(x * 0.9)))
                        b = b.point(lambda x: min(255, int(x * 1.1)))
                        img = Image.merge('RGBA', (r, g, b, a))
                elif filter_type == 'vivid':
                    enhancer = ImageEnhance.Color(img)
                    img = enhancer.enhance(2.0)
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(1.2)
                    enhancer = ImageEnhance.Brightness(img)
                    img = enhancer.enhance(1.1)
                elif filter_type == 'fade':
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    r, g, b, a = img.split()
                    # Reduce opacity to 70% for fade effect
                    a = a.point(lambda x: int(x * 0.7))
                    img = Image.merge('RGBA', (r, g, b, a))
                    # Also reduce color saturation slightly
                    enhancer = ImageEnhance.Color(img)
                    img = enhancer.enhance(0.85)
            except Exception as filter_error:
                logger.warning(f"Failed to apply filter '{filter_type}': {filter_error}")
        
        # Save edited image
        edited_path = os.path.join(get_outputs_dir(), f"edited_{uuid.uuid4().hex}.png")
        img.save(edited_path, 'PNG', optimize=True)
        
        return edited_path
    except Exception as e:
        logger.error(f"Failed to apply image edits: {e}")
        return image_path

# ═══════════════════════════════════════════════════════════════
# CHUNK UPLOAD API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def chunk_upload(request):
    """Handle chunked file uploads for large video files.
    
    - Regular chunks: POST with chunk file + metadata
    - Final assembly: POST with metadata only (no chunk file)
    """
    if request.method != 'POST':
        return _json_error('POST required', 405)

    # Proactively clean old files before accepting new chunks
    try:
        removed, freed_mb = cleanup_old_files(max_age_seconds=600)
        if freed_mb > 0:
            logger.info(f"Pre-cleanup: Removed {removed} files, freed {freed_mb:.1f}MB")
    except Exception as e:
        logger.warning(f"Cleanup before chunk upload failed: {e}")

    upload_id = request.POST.get('upload_id') or str(uuid.uuid4())
    chunk_index = request.POST.get('chunk_index')
    total_chunks = request.POST.get('total_chunks')
    original_filename = sanitize_filename(request.POST.get('filename', 'upload'))
    chunk_file = request.FILES.get('chunk')

    # Log request parameters for debugging
    logger.debug(f"Chunk upload request: upload_id={upload_id}, chunk_index={chunk_index}, total_chunks={total_chunks}, chunk_file={'present' if chunk_file else 'MISSING'}")

    if chunk_index is None or total_chunks is None:
        logger.error(f"Missing parameters: chunk_index={chunk_index}, total_chunks={total_chunks}, POST keys={list(request.POST.keys())}")
        return _json_error('Missing chunk_index or total_chunks', 400)

    try:
        chunk_index = int(chunk_index)
        total_chunks = int(total_chunks)
    except ValueError as e:
        logger.error(f"Invalid chunk numbers: chunk_index={chunk_index}, total_chunks={total_chunks}")
        return _json_error('Invalid chunk numbers', 400)

    # If no chunk file, this is a final assembly request (check if all chunks are present)
    if chunk_file is None:
        logger.debug(f"Final assembly request for upload_id={upload_id}")
    else:
        # Save the chunk file
        try:
            save_chunk(upload_id, chunk_index, chunk_file)
        except OSError as e:
            # Attempt aggressive cleanup if disk is full
            if "No space left" in str(e):
                logger.error(f"Disk space full during chunk upload: {e}")
                try:
                    # Clean files older than 5 minutes
                    cleanup_old_files(max_age_seconds=300)
                    # Retry save after cleanup
                    save_chunk(upload_id, chunk_index, chunk_file)
                except OSError as retry_error:
                    return _json_error('Disk space full: Unable to save chunk', 507)
            else:
                logger.error(f"Error saving chunk: {e}")
                return _json_error(f'Failed to save chunk: {e}', 500)
    
    # Check if assembly was already done (chunks cleaned up but file exists)
    uploads_dir = get_uploads_dir()
    potential_file = os.path.join(uploads_dir, f"{upload_id}{os.path.splitext(original_filename)[1]}")
    
    if os.path.exists(potential_file):
        logger.info(f"Assembly already complete for {upload_id}, returning existing file")
        return JsonResponse({
            'status': 'complete',
            'upload_id': upload_id,
            'file_path': potential_file,
            'filename': original_filename
        })
    
    status = get_chunk_status(upload_id, total_chunks)
    
    logger.info(f"Chunk status: upload_id={upload_id}, received={len(status['received_chunks'])}/{total_chunks}, "
                f"complete={status['complete']}, missing={status['missing_chunks']}")

    if status['complete']:
        try:
            logger.info(f"Assembling {total_chunks} chunks for upload_id={upload_id}")
            final_path = assemble_chunks(upload_id, total_chunks, original_filename)
            logger.info(f"Assembly successful: {final_path}")
            return JsonResponse({
                'status': 'complete',
                'upload_id': upload_id,
                'file_path': final_path,
                'filename': original_filename
            })
        except Exception as e:
            logger.error(f"Assembly failed: {e}", exc_info=True)
            return _json_error(f'Assembly failed: {e}', 500)

    logger.debug(f"Still receiving chunks for {upload_id}: {len(status['received_chunks'])}/{total_chunks}")
    return JsonResponse({
        'status': 'chunk_received',
        'upload_id': upload_id,
        'received': len(status['received_chunks']),
        'total': total_chunks
    })


@csrf_exempt
def chunk_status(request):
    """Check status of a chunked upload."""
    if request.method != 'POST':
        return _json_error('POST required', 405)

    data = json.loads(request.body or '{}')
    upload_id = data.get('upload_id')
    total_chunks = data.get('total_chunks')

    if not upload_id or total_chunks is None:
        return _json_error('Missing upload_id or total_chunks', 400)

    status = get_chunk_status(upload_id, int(total_chunks))
    return JsonResponse(status)

# ═══════════════════════════════════════════════════════════════
# TOOL PAGE VIEWS
# ═══════════════════════════════════════════════════════════════

def _render_tool(request, template, title, description, icon='video'):
    cleanup_old_files()
    return render(request, template, {
        'tool': {'title': title, 'description': description, 'icon': icon}
    })

def converter_page(request):
    return _render_tool(request, 'video_processor/converter.html',
        'Video Converter', 'Convert videos between MP4, AVI, MOV, MKV, and WEBM with quality control.', 'refresh-cw')

def image_to_video_page(request):
    return _render_tool(request, 'video_processor/image_to_video.html',
        'Image to Video', 'Turn your photos into a beautiful slideshow video with transitions.', 'image')

def video_editor_page(request):
    return _render_tool(request, 'video_processor/video_editor.html',
        'Video Editor', 'Trim, cut, rotate, resize, crop, change speed, mute, replace audio, and add text.', 'scissors')

def compressor_page(request):
    return _render_tool(request, 'video_processor/compressor.html',
        'Video Compressor', 'Reduce video file size while maintaining quality with CRF optimization.', 'archive')

def merger_page(request):
    return _render_tool(request, 'video_processor/merger.html',
        'Video Merger', 'Combine multiple videos into one seamless file.', 'combine')

def trimmer_page(request):
    return _render_tool(request, 'video_processor/trimmer.html',
        'Video Trimmer', 'Trim videos to your desired start and end times.', 'crop')

def gif_maker_page(request):
    return _render_tool(request, 'video_processor/gif_maker.html',
        'GIF Maker', 'Convert video clips into animated GIFs with customizable FPS and size.', 'image')

def audio_extractor_page(request):
    return _render_tool(request, 'video_processor/audio_extractor.html',
        'Audio Extractor', 'Extract audio from videos as MP3, WAV, or AAC.', 'music')

def watermark_page(request):
    return _render_tool(request, 'video_processor/watermark.html',
        'Watermark Tool', 'Add text or image watermarks to your videos.', 'stamp')

def subtitle_overlay_page(request):
    return _render_tool(request, 'video_processor/subtitle_overlay.html',
        'Subtitle Overlay', 'Burn subtitles permanently into your videos.', 'subtitles')

# ═══════════════════════════════════════════════════════════════
# MODULE 1: VIDEO CONVERTER API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def convert_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    output_format = request.POST.get('format', 'mp4')
    if not video_file:
        return _json_error('No video uploaded', 400)

    input_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        options = {
            'resolution': request.POST.get('resolution', 'original'),
            'quality': request.POST.get('quality', 'medium'),
            'fps': request.POST.get('fps') or None,
            'codec': request.POST.get('codec') or None,
            'bitrate': request.POST.get('bitrate') or None,
        }
        output_path = _get_output_path(output_format)
        cmd = build_convert_command(input_path, output_path, output_format, options)
        run_ffmpeg(cmd, timeout=1800)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Conversion failed")
        return _json_error(f'Conversion failed: {e}', 500)
    finally:
        safe_remove(input_path)

# ═══════════════════════════════════════════════════════════════
# MODULE 2: IMAGE TO VIDEO API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def image_to_video_api(request):
    """
    Generate video from images with optional audio and per-image edits.
    - Images: uploaded via multipart/form-data
    - image_edits: JSON dict mapping image index to edits {rotation, flipH, flipV, zoom, brightness, contrast, saturation, blur, opacity, filter}
    - image_durations: JSON array of durations per image (seconds)
    - audio: optional audio file
    - audio_trim_start/end: trim audio to this range
    - audio_volume, audio_fade_in, audio_fade_out: audio effects
    - audio_behavior: 'trim' (video to match audio), 'loop' (audio loops), or 'silence'
    - fps, resolution, aspect_ratio, transition: video settings
    """
    if request.method != 'POST':
        return _json_error('POST required', 405)

    images = request.FILES.getlist('images')
    if not images:
        return _json_error('No images uploaded', 400)
    
    if len(images) > 100:
        return _json_error('Maximum 100 images allowed', 400)

    image_paths = []
    output_path = None
    temp_files = []
    audio_path = None
    edited_image_paths = []  # Track edited images for cleanup
    
    try:
        # Upload and validate all images
        for i, img in enumerate(images):
            if i > 99:
                break
            validate_image(img, max_mb=50)
            img_path = save_upload(img)
            image_paths.append(img_path)
            logger.debug(f"Uploaded image {i}: {img.name}")

        if not image_paths:
            return _json_error('Failed to upload images', 400)

        # Parse and apply image edits
        image_edits_raw = request.POST.get('image_edits', '{}')
        try:
            image_edits = json.loads(image_edits_raw) if image_edits_raw else {}
        except (json.JSONDecodeError, ValueError):
            image_edits = {}
            logger.warning("Failed to parse image_edits JSON")

        # Apply edits to each image in order
        for i, img_path in enumerate(image_paths):
            edits = image_edits.get(str(i), {})
            if edits and isinstance(edits, dict):
                try:
                    edited_path = apply_image_edits(img_path, edits)
                    if edited_path != img_path:
                        # Keep track of edited image for cleanup
                        edited_image_paths.append(edited_path)
                        image_paths[i] = edited_path
                        logger.debug(f"Applied edits to image {i}")
                except Exception as e:
                    logger.warning(f"Failed to apply edits to image {i}: {e}")
                    # Continue with original image if edits fail

        # Parse per-image durations
        durations_raw = request.POST.get('image_durations', '[]')
        try:
            durations = json.loads(durations_raw) if durations_raw else []
            # Validate durations
            if not isinstance(durations, list) or len(durations) != len(images):
                durations = [3.0] * len(images)
            # Ensure all durations are positive floats
            durations = [max(0.5, float(d)) for d in durations]
        except (json.JSONDecodeError, ValueError, TypeError):
            durations = [3.0] * len(images)
            logger.warning("Invalid image_durations, using defaults")

        # Parse audio trim settings
        audio_trim_start = 0.0
        audio_trim_end = None
        try:
            audio_trim_start_str = request.POST.get('audio_trim_start', '0')
            audio_trim_start = float(audio_trim_start_str) if audio_trim_start_str else 0.0
            
            audio_trim_end_str = request.POST.get('audio_trim_end', '')
            if audio_trim_end_str:
                audio_trim_end = float(audio_trim_end_str)
        except (ValueError, TypeError):
            audio_trim_start = 0.0
            audio_trim_end = None
            logger.warning("Invalid audio trim parameters")

        # Parse audio effect settings
        audio_volume = 1.0
        audio_fade_in = 0.0
        audio_fade_out = 0.0
        try:
            audio_volume_str = request.POST.get('audio_volume', '1.0')
            audio_volume = float(audio_volume_str) if audio_volume_str else 1.0
            
            audio_fade_in_str = request.POST.get('audio_fade_in', '0')
            audio_fade_in = float(audio_fade_in_str) if audio_fade_in_str else 0.0
            
            audio_fade_out_str = request.POST.get('audio_fade_out', '0')
            audio_fade_out = float(audio_fade_out_str) if audio_fade_out_str else 0.0
        except (ValueError, TypeError):
            logger.warning("Invalid audio effect parameters")

        # Sanitize filename for download
        original_filename = request.POST.get('filename', 'slideshow.mp4')
        if not isinstance(original_filename, str):
            original_filename = 'slideshow.mp4'
        original_filename = sanitize_filename(original_filename)
        if not original_filename.endswith('.mp4'):
            original_filename += '.mp4'

        # Build options dict for FFmpeg builder
        options = {
            'durations': durations,
            'fps': max(15, min(60, int(request.POST.get('fps', 30)))),
            'resolution': request.POST.get('resolution', '1080p'),
            'aspect_ratio': request.POST.get('aspect_ratio', '16:9'),
            'transition': request.POST.get('transition', 'none'),
            'audio_trim_start': audio_trim_start,
            'audio_trim_end': audio_trim_end,
            'audio_volume': max(0.0, min(2.0, audio_volume)),
            'audio_fade_in': max(0.0, audio_fade_in),
            'audio_fade_out': max(0.0, audio_fade_out),
            'audio_behavior': request.POST.get('audio_behavior', 'trim'),  # trim, loop, silence
            'filename': original_filename,
        }

        # Handle optional audio file
        audio_file = request.FILES.get('audio')
        if audio_file:
            try:
                validate_audio(audio_file, max_mb=500)
                audio_path = save_upload(audio_file)
                options['audio_path'] = audio_path
                logger.debug(f"Uploaded audio: {audio_file.name}")
            except ValidationError as e:
                logger.warning(f"Audio validation failed: {e}")
                # Continue without audio if validation fails
            except OSError as e:
                if "No space left" in str(e):
                    return _json_error('Disk space full: Cannot upload audio file', 507)
                raise

        # Generate video
        output_path = _get_output_path('mp4')
        output_path, temp_files = build_image_to_video_command(image_paths, output_path, options)
        
        logger.info(f"Generated video: {output_path}")
        return _cleanup_response(output_path, original_filename)
        
    except ValidationError as e:
        logger.warning(f"Validation error: {e}")
        return _json_error(str(e), 400)
    except OSError as e:
        if "No space left" in str(e):
            logger.error(f"Disk space full: {e}")
            return _json_error('Disk space full: Unable to process request', 507)
        logger.exception("OS error during processing")
        return _json_error(f'File operation failed: {str(e)[:100]}', 500)
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Image to video generation failed")
        return _json_error(f'Processing failed: {str(e)[:100]}', 500)
    finally:
        # Comprehensive cleanup of all temporary files
        try:
            for p in image_paths:
                if p:
                    safe_remove(p)
            for tf in temp_files:
                if tf:
                    safe_remove(tf)
            for edited_path in edited_image_paths:
                if edited_path:
                    safe_remove(edited_path)
            if audio_path:
                safe_remove(audio_path)
            logger.debug("Temp files cleaned up after image-to-video generation")
        except Exception as cleanup_error:
            logger.error(f"Cleanup error: {cleanup_error}")

# ═══════════════════════════════════════════════════════════════
# MODULE 3: VIDEO EDITOR APIs
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def trim_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    use_chunk_path = request.POST.get('use_chunk_path')
    upload_id = request.POST.get('upload_id')
    filename = request.POST.get('filename')

    input_path = output_path = None
    video_file_name = None
    
    try:
        if use_chunk_path and upload_id and filename:
            ext = os.path.splitext(filename)[1]
            input_path = os.path.join(get_uploads_dir(), f"{upload_id}{ext}")
            if not os.path.exists(input_path):
                return _json_error(f'Assembled file not found: {input_path}', 404)
            video_file_name = filename
            logger.info(f"Using chunk-assembled file: {input_path}")
        else:
            input_path = _handle_uploaded_video(request, 'video')
            video_file = request.FILES.get('video')
            video_file_name = video_file.name if video_file else 'trimmed.mp4'

        start = float(request.POST.get('start', 0))
        end = float(request.POST.get('end', 0))
        output_path = _get_output_path('mp4')
        cmd = build_trim_command(input_path, output_path, start, end)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file_name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Trim failed")
        return _json_error(f'Trim failed: {e}', 500)
    finally:
        if not use_chunk_path:
            safe_remove(input_path)


@csrf_exempt
def cut_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = output_path = None
    temp_files = []
    try:
        input_path = _handle_uploaded_video(request, 'video')
        segments_raw = request.POST.get('segments', '[]')
        segments = json.loads(segments_raw)
        if not segments:
            return _json_error('No cut segments provided', 400)

        output_path = _get_output_path('mp4')
        cmd, temp_files = build_cut_command(input_path, output_path, [
            (float(s['start']), float(s['end'])) for s in segments
        ])
        run_ffmpeg(cmd, timeout=900)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Cut failed")
        return _json_error(f'Cut failed: {e}', 500)
    finally:
        safe_remove(input_path)
        for tf in temp_files:
            safe_remove(tf)


@csrf_exempt
def rotate_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        angle = request.POST.get('angle', '90')
        output_path = _get_output_path('mp4')
        cmd = build_rotate_command(input_path, output_path, angle)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Rotate failed")
        return _json_error(f'Rotate failed: {e}', 500)
    finally:
        safe_remove(input_path)


@csrf_exempt
def resize_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        width = int(request.POST.get('width', 1280))
        height = int(request.POST.get('height', 720))
        output_path = _get_output_path('mp4')
        cmd = build_resize_command(input_path, output_path, width, height)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Resize failed")
        return _json_error(f'Resize failed: {e}', 500)
    finally:
        safe_remove(input_path)


@csrf_exempt
def crop_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        x = int(request.POST.get('x', 0))
        y = int(request.POST.get('y', 0))
        width = int(request.POST.get('width', 100))
        height = int(request.POST.get('height', 100))
        output_path = _get_output_path('mp4')
        cmd = build_crop_command(input_path, output_path, x, y, width, height)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Crop failed")
        return _json_error(f'Crop failed: {e}', 500)
    finally:
        safe_remove(input_path)


@csrf_exempt
def speed_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        speed = float(request.POST.get('speed', 1.0))
        output_path = _get_output_path('mp4')
        cmd = build_speed_command(input_path, output_path, speed)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Speed change failed")
        return _json_error(f'Speed change failed: {e}', 500)
    finally:
        safe_remove(input_path)


@csrf_exempt
def mute_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        output_path = _get_output_path('mp4')
        cmd = build_mute_command(input_path, output_path)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Mute failed")
        return _json_error(f'Mute failed: {e}', 500)
    finally:
        safe_remove(input_path)


@csrf_exempt
def replace_audio_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = audio_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        audio_file = request.FILES.get('audio')
        if not audio_file:
            return _json_error('No audio file uploaded', 400)
        audio_path = _handle_uploaded_audio(request, 'audio')
        output_path = _get_output_path('mp4')
        cmd = build_replace_audio_command(input_path, audio_path, output_path)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Replace audio failed")
        return _json_error(f'Replace audio failed: {e}', 500)
    finally:
        safe_remove(input_path)
        safe_remove(audio_path)


@csrf_exempt
def text_overlay_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        text = request.POST.get('text', '')
        if not text:
            return _json_error('Text is required', 400)
        options = {
            'position': request.POST.get('position', 'center'),
            'font_size': int(request.POST.get('font_size', 36)),
            'color': request.POST.get('color', 'white'),
            'start': float(request.POST.get('start', 0)),
            'duration': float(request.POST.get('duration', 5)),
        }
        output_path = _get_output_path('mp4')
        cmd = build_text_overlay_command(input_path, output_path, text, options)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Text overlay failed")
        return _json_error(f'Text overlay failed: {e}', 500)
    finally:
        safe_remove(input_path)

# ═══════════════════════════════════════════════════════════════
# MODULE 4: VIDEO COMPRESSOR API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def compress_video_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    # Check if using chunk assembly path
    use_chunk_path = request.POST.get('use_chunk_path')
    upload_id = request.POST.get('upload_id')
    filename = request.POST.get('filename')
    
    input_path = output_path = None
    video_file_name = None
    
    try:
        if use_chunk_path and upload_id and filename:
            # File was assembled from chunks
            ext = os.path.splitext(filename)[1]
            input_path = os.path.join(get_uploads_dir(), f"{upload_id}{ext}")
            if not os.path.exists(input_path):
                return _json_error(f'Assembled file not found: {input_path}', 404)
            video_file_name = filename
            logger.info(f"Using chunk-assembled file: {input_path}")
        else:
            # Regular file upload
            input_path = _handle_uploaded_video(request, 'video')
            video_file_name = request.FILES.get('video').name if request.FILES.get('video') else 'video.mp4'
        
        options = {
            'quality': request.POST.get('quality', 'medium'),
            'codec': request.POST.get('codec', 'libx264'),
            'target_size_mb': request.POST.get('target_size_mb') or None,
        }
        if options['target_size_mb']:
            options['target_size_mb'] = float(options['target_size_mb'])
        
        output_path = _get_output_path('mp4')
        
        # Try async with Celery, fall back to background thread if Redis unavailable
        try:
            from .tasks import compress_video_task
            task = compress_video_task.delay(input_path, output_path, options)
            logger.info(f"Started async compression task: {task.id}")
            return JsonResponse({
                'status': 'processing',
                'task_id': task.id,
                'output_filename': video_file_name,
                'is_celery': True,
                'message': 'Compression started. Checking status...'
            })
        except Exception as celery_error:
            logger.warning(f"Celery unavailable, falling back to background thread: {celery_error}")
            # Fall back to background thread compression
            job_id = str(uuid.uuid4())
            _background_jobs[job_id] = {'status': 'processing'}
            
            thread = threading.Thread(
                target=_run_compression_background,
                args=(input_path, output_path, options, job_id, use_chunk_path),
                daemon=True
            )
            thread.start()
            
            return JsonResponse({
                'status': 'processing',
                'task_id': job_id,
                'output_filename': video_file_name,
                'is_celery': False,
                'message': 'Compression started in background. Checking status...'
            })
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Compression failed")
        return _json_error(f'Compression failed: {e}', 500)


@csrf_exempt
def compress_task_status(request):
    """Check status of compression (Celery task or background thread) and download when ready."""
    if request.method != 'GET':
        return _json_error('GET required', 405)
    
    task_id = request.GET.get('task_id')
    if not task_id:
        return _json_error('task_id required', 400)
    
    try:
        # Check if it's a background job first
        if task_id in _background_jobs:
            job = _background_jobs[task_id]
            
            if job['status'] == 'processing':
                return JsonResponse({
                    'status': 'processing',
                    'message': 'Compressing video...'
                })
            elif job['status'] == 'complete':
                output_path = job['output_path']
                output_filename = request.GET.get('filename', 'compressed.mp4')
                
                if os.path.exists(output_path):
                    logger.info(f"Background job {task_id} completed. Serving file: {output_path}")
                    return _cleanup_response(output_path, output_filename)
                else:
                    return _json_error(f'Output file not found: {output_path}', 404)
            elif job['status'] == 'failed':
                logger.error(f"Background job {task_id} failed: {job.get('error')}")
                return _json_error(f"Compression failed: {job.get('error')}", 500)
        
        # Otherwise check if it's a Celery task
        from celery.result import AsyncResult
        task_result = AsyncResult(task_id)
        
        if task_result.state == 'PENDING':
            return JsonResponse({
                'status': 'pending',
                'message': 'Compression starting...'
            })
        elif task_result.state == 'PROGRESS':
            return JsonResponse({
                'status': 'processing',
                'message': 'Compressing video...',
                'info': task_result.info
            })
        elif task_result.state == 'SUCCESS':
            result = task_result.result
            output_path = result.get('output_path')
            output_filename = request.GET.get('filename', 'compressed.mp4')
            
            if os.path.exists(output_path):
                logger.info(f"Celery task {task_id} completed. Serving file: {output_path}")
                return _cleanup_response(output_path, output_filename)
            else:
                return _json_error(f'Output file not found: {output_path}', 404)
        elif task_result.state == 'FAILURE':
            logger.error(f"Celery task {task_id} failed: {task_result.info}")
            return _json_error(f'Compression failed: {task_result.info}', 500)
        else:
            return JsonResponse({
                'status': task_result.state.lower(),
                'message': 'Processing...'
            })
    except Exception as e:
        logger.exception("Status check failed")
        return _json_error(f'Status check failed: {e}', 500)

# ═══════════════════════════════════════════════════════════════
# MODULE 5: VIDEO MERGER API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def merge_videos_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_files = request.FILES.getlist('videos')
    if len(video_files) < 2:
        return _json_error('Upload at least 2 videos', 400)

    video_paths = []
    output_path = None
    concat_file = None
    try:
        for vf in video_files:
            validate_video(vf)
            video_paths.append(save_upload(vf))

        output_path = _get_output_path('mp4')
        cmd, concat_file = build_merge_command(video_paths, output_path)
        run_ffmpeg(cmd, timeout=1800)
        return _cleanup_response(output_path, video_files[0].name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Merge failed")
        return _json_error(f'Merge failed: {e}', 500)
    finally:
        for vp in video_paths:
            safe_remove(vp)
        safe_remove(concat_file)

# ═══════════════════════════════════════════════════════════════
# MODULE 6: GIF MAKER API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def make_gif_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    use_chunk_path = request.POST.get('use_chunk_path')
    upload_id = request.POST.get('upload_id')
    filename = request.POST.get('filename')

    input_path = output_path = None
    video_file_name = None
    
    try:
        if use_chunk_path and upload_id and filename:
            ext = os.path.splitext(filename)[1]
            input_path = os.path.join(get_uploads_dir(), f"{upload_id}{ext}")
            if not os.path.exists(input_path):
                return _json_error(f'Assembled file not found: {input_path}', 404)
            video_file_name = filename
            logger.info(f"Using chunk-assembled file: {input_path}")
        else:
            input_path = _handle_uploaded_video(request, 'video')
            video_file = request.FILES.get('video')
            video_file_name = video_file.name if video_file else 'output.gif'

        options = {
            'fps': int(request.POST.get('fps', 10)),
            'width': int(request.POST.get('width', 480)),
            'start': float(request.POST.get('start', 0)),
            'duration': float(request.POST.get('duration', 5)),
        }
        output_path = _get_output_path('gif')
        cmd = build_gif_command(input_path, output_path, options)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file_name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("GIF creation failed")
        return _json_error(f'GIF creation failed: {e}', 500)
    finally:
        if not use_chunk_path:
            safe_remove(input_path)

# ═══════════════════════════════════════════════════════════════
# MODULE 7: AUDIO EXTRACTOR API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def extract_audio_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    use_chunk_path = request.POST.get('use_chunk_path')
    upload_id = request.POST.get('upload_id')
    filename = request.POST.get('filename')

    input_path = output_path = None
    video_file_name = None

    try:
        if use_chunk_path and upload_id and filename:
            ext = os.path.splitext(filename)[1]
            input_path = os.path.join(get_uploads_dir(), f"{upload_id}{ext}")
            if not os.path.exists(input_path):
                return _json_error(f'Assembled file not found: {input_path}', 404)
            video_file_name = filename
            logger.info(f"Using chunk-assembled file: {input_path}")
        else:
            input_path = _handle_uploaded_video(request, 'video')
            video_file = request.FILES.get('video')
            video_file_name = video_file.name if video_file else 'output.audio'

        fmt = request.POST.get('format', 'mp3')
        options = {
            'format': fmt,
            'quality': request.POST.get('quality', '192k'),
            'start': float(request.POST.get('start', 0)),
            'duration': float(request.POST.get('duration', 0)),
            'sample_rate': request.POST.get('sample_rate'),
            'channels': request.POST.get('channels'),
        }

        output_path = _get_output_path(fmt)
        cmd = build_audio_extract_command(input_path, output_path, options)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file_name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        error_msg = str(e)
        if "Output file does not contain any stream" in error_msg:
            logger.warning("Attempted to extract audio from a video with no audio track.")
            return _json_error('This video does not contain an audio track to extract.', 400)
        logger.exception("Audio extraction failed")
        return _json_error(f'Audio extraction failed: {e}', 500)
    finally:
        if not use_chunk_path:
            safe_remove(input_path)

# ═══════════════════════════════════════════════════════════════
# MODULE 8: WATERMARK API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def add_watermark_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = output_path = None
    image_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        options = {
            'text': request.POST.get('text', ''),
            'position': request.POST.get('position', 'bottom-right'),
            'font_size': int(request.POST.get('font_size', 24)),
            'color': request.POST.get('color', 'white'),
            'opacity': float(request.POST.get('opacity', 0.7)),
        }
        image_file = request.FILES.get('watermark_image')
        if image_file:
            image_path = save_upload(image_file)
            options['image_path'] = image_path

        output_path = _get_output_path('mp4')
        cmd = build_watermark_command(input_path, output_path, options)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Watermark failed")
        return _json_error(f'Watermark failed: {e}', 500)
    finally:
        safe_remove(input_path)
        safe_remove(image_path)

# ═══════════════════════════════════════════════════════════════
# MODULE 9: SUBTITLE OVERLAY API
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def add_subtitle_api(request):
    if request.method != 'POST':
        return _json_error('POST required', 405)

    video_file = request.FILES.get('video')
    input_path = subtitle_path = output_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video')
        subtitle_file = request.FILES.get('subtitle')
        if not subtitle_file:
            return _json_error('No subtitle file uploaded', 400)
        subtitle_path = _handle_uploaded_subtitle(request, 'subtitle')

        options = {
            'font_size': int(request.POST.get('font_size', 24)),
            'color': request.POST.get('color', 'white'),
            'outline': int(request.POST.get('outline', 1)),
            'bold': int(request.POST.get('bold', 0)),
        }
        output_path = _get_output_path('mp4')
        cmd = build_subtitle_command(input_path, subtitle_path, output_path, options)
        run_ffmpeg(cmd, timeout=600)
        return _cleanup_response(output_path, video_file.name)
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        safe_remove(output_path)
        logger.exception("Subtitle overlay failed")
        return _json_error(f'Subtitle overlay failed: {e}', 500)
    finally:
        safe_remove(input_path)
        safe_remove(subtitle_path)

# ═══════════════════════════════════════════════════════════════
# MODULE 10: VIDEO INFO API (shared utility for frontend)
# ═══════════════════════════════════════════════════════════════

@csrf_exempt
def video_info_api(request):
    """Return video metadata (duration, resolution, fps)."""
    if request.method != 'POST':
        return _json_error('POST required', 405)

    input_path = None
    try:
        input_path = _handle_uploaded_video(request, 'video', max_mb=50)
        info = get_video_info(input_path)
        return JsonResponse({
            'status': 'success',
            'info': info
        })
    except ValidationError as e:
        return _json_error(str(e))
    except Exception as e:
        logger.exception("Video info failed")
        return _json_error(f'Failed: {e}', 500)
    finally:
        safe_remove(input_path)

 