import os
import mimetypes
from pathlib import Path
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from django import forms

from .utils import (
    save_uploaded_file,
    blur_image,
    brighten_image,
    change_image_background,
    remove_image_background,
    compress_image,
    resize_image,
    rotate_image,
    watermark_image,
    crop_image,
    merge_images,
    change_gif_speed,
    extract_image_from_video,
    gif_to_video,
    image_to_video,
    convert_image,
    format_download_name
)

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

# ─── New Image Tools Configuration ────────────────────────────────────────
IMAGE_TOOLS = {
    'blur-image': {
        'title': 'Blur Image',
        'description': 'Add a professional blur effect to your images with custom radius.',
        'icon': 'cloud-fog',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#8b5cf6',
        'gradient': 'from-purple-500 to-indigo-600',
    },
    'brighten-image': {
        'title': 'Brighten Image',
        'description': 'Adjust the brightness of your images to make them pop.',
        'icon': 'sun',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#facc15',
        'gradient': 'from-yellow-400 to-amber-500',
    },
    'change-gif-speed': {
        'title': 'Change GIF Speed',
        'description': 'Speed up or slow down your animated GIFs easily.',
        'icon': 'zap',
        'accept': '.gif',
        'allowed_extensions': ['.gif'],
        'category': 'image-pro',
        'color': '#ef4444',
        'gradient': 'from-red-500 to-orange-600',
    },
    'change-background': {
        'title': 'Change Background',
        'description': 'Remove subject and replace it with a custom solid background color.',
        'icon': 'palette',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#10b981',
        'gradient': 'from-emerald-400 to-teal-500',
    },
    'compress-image': {
        'title': 'Compress Image',
        'description': 'Reduce image file size with optimal quality compression.',
        'icon': 'archive',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#dc2626',
        'gradient': 'from-red-500 to-red-700',
    },
    'cut-image': {
        'title': 'Cut Image',
        'description': 'Crop your image to focus on what matters.',
        'icon': 'crop',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#84cc16',
        'gradient': 'from-lime-500 to-lime-700',
    },
    'extract-frame': {
        'title': 'Extract Frame',
        'description': 'Save a specific frame from a video as a high-quality image.',
        'icon': 'video',
        'accept': '.mp4,.mov,.avi',
        'allowed_extensions': ['.mp4', '.mov', '.avi'],
        'category': 'image-pro',
        'color': '#0ea5e9',
        'gradient': 'from-sky-400 to-blue-600',
    },
    'gif-to-video': {
        'title': 'GIF to Video',
        'description': 'Convert animated GIFs into standard MP4 or WEBM video files with advanced effects and music.',
        'icon': 'film',
        'accept': '.gif',
        'allowed_extensions': ['.gif'],
        'category': 'image-pro',
        'color': '#6366f1',
        'gradient': 'from-indigo-500 to-violet-600',
    },
    'image-to-video': {
        'title': 'Image to Video',
        'description': 'Create professional videos from your images with music, transitions, and custom durations.',
        'icon': 'monitor-play',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#7c3aed',
        'gradient': 'from-violet-500 to-purple-700',
        'multi_file': True,
    },
    'merge-images': {
        'title': 'Merge Images',
        'description': 'Combine multiple images side-by-side or stacked.',
        'icon': 'combine',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#f43f5e',
        'gradient': 'from-rose-500 to-pink-600',
        'multi_file': True,
    },
    'remove-background': {
        'title': 'Remove Background',
        'description': 'Automatically remove image backgrounds with AI precision.',
        'icon': 'eraser',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#06b6d4',
        'gradient': 'from-cyan-400 to-blue-500',
    },
    'resize-image': {
        'title': 'Resize Image',
        'description': 'Resize your image by width, height, or percentage.',
        'icon': 'move',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#f97316',
        'gradient': 'from-orange-500 to-red-500',
    },
    'rotate-image': {
        'title': 'Rotate Image',
        'description': 'Rotate your images clockwise or counter-clockwise.',
        'icon': 'rotate-cw',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#8b5cf6',
        'gradient': 'from-purple-500 to-indigo-600',
    },
    'watermark-image': {
        'title': 'Watermark Image',
        'description': 'Protect your brand by adding custom text watermarks to your photos.',
        'icon': 'stamp',
        'accept': '.jpg,.jpeg,.png',
        'allowed_extensions': ['.jpg', '.jpeg', '.png'],
        'category': 'image-pro',
        'color': '#0891b2',
        'gradient': 'from-cyan-500 to-cyan-700',
    },

    # --- Converters ---
    'image-converter': {
        'title': 'Image Converter',
        'description': 'Convert images between multiple formats like JPG, PNG, WEBP, etc.',
        'icon': 'refresh-cw',
        'accept': '.jpg,.jpeg,.png,.bmp,.webp',
        'allowed_extensions': ['.jpg', '.jpeg', '.png', '.bmp', '.webp'],
        'category': 'image-conv',
        'color': '#475569',
        'gradient': 'from-slate-500 to-slate-700',
    },
    'jpg-converter': { 'title': 'JPG Converter', 'description': 'Convert any image format to JPG.', 'icon': 'file-image', 'accept': '.png,.bmp,.webp,.tiff', 'allowed_extensions': ['.png', '.bmp', '.webp', '.tiff'], 'category': 'image-conv', 'color': '#2b6cb0', 'gradient': 'from-blue-500 to-blue-700', 'target': 'jpg' },
    'png-converter': { 'title': 'PNG Converter', 'description': 'Convert any image format to PNG.', 'icon': 'file-image', 'accept': '.jpg,.jpeg,.bmp,.webp,.tiff', 'allowed_extensions': ['.jpg', '.jpeg', '.bmp', '.webp', '.tiff'], 'category': 'image-conv', 'color': '#276749', 'gradient': 'from-green-500 to-emerald-700', 'target': 'png' },
    'bmp-converter': { 'title': 'BMP Converter', 'description': 'Convert any image format to Windows Bitmap.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png', '.webp'], 'category': 'image-conv', 'color': '#c05621', 'gradient': 'from-orange-500 to-red-500', 'target': 'bmp' },
    'gif-converter': { 'title': 'GIF Converter', 'description': 'Convert static images to GIF format.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#6b46c1', 'gradient': 'from-purple-500 to-indigo-700', 'target': 'gif' },
    'pdf-converter': { 'title': 'PDF Converter', 'description': 'Convert your images directly into a PDF document.', 'icon': 'file-pdf', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#dc2626', 'gradient': 'from-red-500 to-red-700', 'target': 'pdf' },
    'tiff-converter': { 'title': 'TIFF Converter', 'description': 'High-quality TIFF conversion for professional printing.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#0d9488', 'gradient': 'from-teal-500 to-teal-700', 'target': 'tiff' },
    'webp-converter': { 'title': 'WEBP Converter', 'description': 'Optimize your images for the web with WEBP format.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#0ea5e9', 'gradient': 'from-sky-500 to-sky-700', 'target': 'webp' },
    'dng-converter': { 'title': 'DNG Converter', 'description': 'DNG Digital Negative conversion placeholder.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#111827', 'gradient': 'from-gray-700 to-black', 'target': 'tiff' },
}

def tool_page(request, tool_slug):
    if tool_slug not in IMAGE_TOOLS:
        raise Http404("Tool not found")
    
    tool = IMAGE_TOOLS[tool_slug]
    
    # Selecting the template. Use a generic converter template for now if it exists,
    # or create a custom one for specific parameter needs.
    template = 'image_processor/tool_detail.html'
    
    context = {
        'tool': tool,
        'tool_slug': tool_slug,
        'page_title': f'{tool["title"]} — Image Editor',
    }
    return render(request, template, context)

@csrf_exempt
@require_POST
def process_tool(request, tool_slug):
    if tool_slug not in IMAGE_TOOLS:
        return JsonResponse({'error': 'Tool not found'}, status=404)
    
    # Generic handle for multi-file vs single file
    if IMAGE_TOOLS[tool_slug].get('multi_file'):
        files = request.FILES.getlist('files')
        if not files: return JsonResponse({'error': 'No files uploaded'}, status=400)
        input_paths = [save_uploaded_file(f) for f in files]
        original_name = files[0].name
    else:
        uploaded_file = request.FILES.get('file')
        if not uploaded_file: return JsonResponse({'error': 'No file uploaded'}, status=400)
        input_path = save_uploaded_file(uploaded_file)
        input_paths = [input_path]
        original_name = uploaded_file.name

    try:
        output_path = None
        
        if tool_slug == 'blur-image':
            radius = request.POST.get('radius', 5)
            output_path = blur_image(input_paths[0], original_name, radius=int(radius))
        elif tool_slug == 'brighten-image':
            factor = request.POST.get('factor', 1.5)
            output_path = brighten_image(input_paths[0], original_name, factor=float(factor))
        elif tool_slug == 'change-background':
            hex_color = request.POST.get('color', '#ffffff').lstrip('#')
            bg_color = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            output_path = change_image_background(input_paths[0], original_name, bg_color=bg_color)
        elif tool_slug == 'remove-background':
            output_path = remove_image_background(input_paths[0], original_name)
        elif tool_slug == 'compress-image':
            quality = request.POST.get('quality', 30)
            output_path = compress_image(input_paths[0], original_name, quality=int(quality))
        elif tool_slug == 'resize-image':
            width = request.POST.get('width')
            height = request.POST.get('height')
            output_path = resize_image(input_paths[0], original_name, width=width, height=height)
        elif tool_slug == 'rotate-image':
            angle = request.POST.get('angle', 90)
            output_path = rotate_image(input_paths[0], original_name, angle=angle)
        elif tool_slug == 'watermark-image':
            text = request.POST.get('text', 'ScanPDF')
            output_path = watermark_image(input_paths[0], original_name, text=text)
        elif tool_slug == 'cut-image':
            l, t, r, b = request.POST.get('left'), request.POST.get('top'), request.POST.get('right'), request.POST.get('bottom')
            output_path = crop_image(input_paths[0], original_name, l, t, r, b)
        elif tool_slug == 'merge-images':
            direction = request.POST.get('direction', 'horizontal')
            output_path = merge_images(input_paths, original_name, direction=direction)
        elif tool_slug == 'change-gif-speed':
            factor = request.POST.get('speed', 1.0)
            output_path = change_gif_speed(input_paths[0], original_name, speed_factor=factor)
        elif tool_slug == 'extract-frame':
            ts = request.POST.get('timestamp', 1.0)
            output_path = extract_image_from_video(input_paths[0], original_name, timestamp=ts)
        elif tool_slug == 'gif-to-video':
            target_format = request.POST.get('target_format', 'mp4')
            duration = request.POST.get('duration', 'default')
            effect = request.POST.get('effect', 'none')
            effect_duration = request.POST.get('effect_duration', '3')
            speed = request.POST.get('speed', 'default')
            color = request.POST.get('color', 'default')
            
            music_file = request.FILES.get('music_file')
            music_path = save_uploaded_file(music_file) if music_file else None
            
            output_path = gif_to_video(
                input_paths[0], 
                original_name, 
                target_format=target_format,
                duration=duration,
                effect=effect,
                effect_duration=effect_duration,
                speed_factor=speed,
                bg_color=color,
                music_path=music_path
            )
            if music_path and os.path.exists(music_path):
                input_paths.append(music_path)
        elif tool_slug == 'image-to-video':
            target_format = request.POST.get('target_format', 'mp4')
            duration_per_image = request.POST.get('img_duration', '2')
            transition = request.POST.get('transition', 'fade')
            
            music_file = request.FILES.get('music_file')
            music_path = save_uploaded_file(music_file) if music_file else None
            
            output_path = image_to_video(
                input_paths, 
                original_name, 
                target_format=target_format,
                duration_per_image=duration_per_image,
                transition_type=transition,
                music_path=music_path
            )
            if music_path and os.path.exists(music_path):
                input_paths.append(music_path)
        
        # --- Converters ---
        elif tool_slug.endswith('-converter'):
            target = request.POST.get('target_format') or IMAGE_TOOLS[tool_slug].get('target', 'jpg')
            output_path = convert_image(input_paths[0], original_name, target)
        
        if output_path and os.path.exists(output_path):
            return create_cleanup_response(output_path)
        else:
            return JsonResponse({'error': 'Failed to process file.'}, status=500)

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    finally:
        # Clean up input files
        for p in input_paths:
            if os.path.exists(p):
                try: os.remove(p)
                except: pass
