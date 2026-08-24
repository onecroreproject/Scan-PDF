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

        's1': 'Blur Images',
        'highlight': 'Online in Seconds',
        'seo_intro': 'Blur images online with our easy-to-use image blur tool. Apply a smooth blur effect to JPG, JPEG, and PNG images without installing software.',
        'seo_keywords': 'blur image online, blur photo, image blur tool, blur JPG, blur PNG, online image editor',
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

        's1': 'Brighten Images',
        'highlight': 'Online Easily',
        'seo_intro': 'Brighten images online quickly and easily. Improve dark photos and adjust image brightness without complicated editing software.',
        'seo_keywords': 'brighten image online, brighten photo, image brightness tool, brighten JPG, brighten PNG',
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

        's1': 'Change GIF Speed',
        'highlight': 'Online Instantly',
        'seo_intro': 'Change GIF animation speed online. Speed up or slow down animated GIFs easily without installing any software.',
        'seo_keywords': 'change GIF speed, GIF speed changer, speed up GIF, slow down GIF, animated GIF editor',
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

        's1': 'Change Image',
        'highlight': 'Background Online',
        'seo_intro': 'Change image backgrounds online with an easy background editing tool. Replace your existing background with a custom color.',
        'seo_keywords': 'change image background, background changer, replace image background, photo background editor, online background changer',
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

        's1': 'Compress Images',
        'highlight': 'Without Losing Quality',
        'seo_intro': 'Compress JPG, JPEG, and PNG images online to reduce file size while maintaining excellent image quality.',
        'seo_keywords': 'compress image online, image compressor, compress JPG, compress PNG, reduce image size, photo compressor',
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

        's1': 'Crop Images',
        'highlight': 'Online with Ease',
        'seo_intro': 'Crop images online quickly and easily. Select the area you want to keep and create the perfect image dimensions.',
        'seo_keywords': 'crop image online, crop photo, image cropper, JPG cropper, PNG cropper, online photo editor',
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

        's1': 'Merge Images',
        'highlight': 'Online into One',
        'seo_intro': 'Merge multiple images into one image online. Combine JPG and PNG files horizontally or vertically with ease.',
        'seo_keywords': 'merge images online, combine images, join photos, merge JPG, merge PNG, image merger',
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

        's1': 'Remove Image',
        'highlight': 'Background Automatically',
        'seo_intro': 'Remove image backgrounds online with AI-powered background removal. Create clean transparent images quickly and easily.',
        'seo_keywords': 'remove background online, background remover, remove photo background, transparent background, AI background remover',
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

        's1': 'Resize Images',
        'highlight': 'Online to Any Size',
        'seo_intro': 'Resize JPG, JPEG, and PNG images online by setting custom width and height. Quickly optimize images for websites and social media.',
        'seo_keywords': 'resize image online, image resizer, resize photo, resize JPG, resize PNG, image dimensions',
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

        's1': 'Rotate Images',
        'highlight': 'Online in One Click',
        'seo_intro': 'Rotate images online clockwise or counter-clockwise. Quickly fix image orientation without installing an image editor.',
        'seo_keywords': 'rotate image online, rotate photo, image rotator, rotate JPG, rotate PNG, photo rotation tool',
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

        's1': 'Watermark Images',
        'highlight': 'Online Easily',
        'seo_intro': 'Add text watermarks to images online to protect your photos and brand. Create customized watermarked JPG and PNG images.',
        'seo_keywords': 'watermark image online, add watermark, photo watermark, image watermark tool, text watermark, watermark JPG',
    },

    'image-converter': {
        'title': 'Image Converter',
        'description': 'Convert images between multiple formats like JPG, PNG, WEBP, etc.',
        'icon': 'refresh-cw',
        'accept': '.jpg,.jpeg,.png,.bmp,.webp',
        'allowed_extensions': ['.jpg', '.jpeg', '.png', '.bmp', '.webp'],
        'category': 'image-conv',
        'color': '#475569',
        'gradient': 'from-slate-500 to-slate-700',

        's1': 'Convert Images',
        'highlight': 'Between Any Format',
        'seo_intro': 'Convert images online between popular formats including JPG, PNG, BMP, WEBP, and more with our simple image converter.',
        'seo_keywords': 'image converter, convert image online, JPG converter, PNG converter, WEBP converter, image format converter',
    },

    'jpg-converter': {
        'title': 'JPG Converter',
        'description': 'Convert any image format to JPG.',
        'icon': 'file-image',
        'accept': '.png,.bmp,.webp,.tiff',
        'allowed_extensions': ['.png', '.bmp', '.webp', '.tiff'],
        'category': 'image-conv',
        'color': '#2b6cb0',
        'gradient': 'from-blue-500 to-blue-700',
        'target': 'jpg',

        's1': 'Convert Images to',
        'highlight': 'JPG Online',
        'seo_intro': 'Convert PNG, BMP, WEBP, and TIFF images to JPG online quickly. Get high-quality JPG files with our free image converter.',
        'seo_keywords': 'JPG converter, convert to JPG, PNG to JPG, WEBP to JPG, BMP to JPG, TIFF to JPG',
    },

    'png-converter': {
        'title': 'PNG Converter',
        'description': 'Convert any image format to PNG.',
        'icon': 'file-image',
        'accept': '.jpg,.jpeg,.bmp,.webp,.tiff',
        'allowed_extensions': ['.jpg', '.jpeg', '.bmp', '.webp', '.tiff'],
        'category': 'image-conv',
        'color': '#276749',
        'gradient': 'from-green-500 to-emerald-700',
        'target': 'png',

        's1': 'Convert Images to',
        'highlight': 'PNG Online',
        'seo_intro': 'Convert JPG, JPEG, BMP, WEBP, and TIFF images to PNG online while preserving excellent image quality.',
        'seo_keywords': 'PNG converter, convert to PNG, JPG to PNG, WEBP to PNG, BMP to PNG, image converter',
    },

    'jpg-converter': { 'title': 'JPG Converter', 'description': 'Convert any image format to JPG.', 'icon': 'file-image', 'accept': '.png,.bmp,.webp,.tiff', 'allowed_extensions': ['.png', '.bmp', '.webp', '.tiff'], 'category': 'image-conv', 'color': '#2b6cb0', 'gradient': 'from-blue-500 to-blue-700', 'target': 'jpg' },
    'png-converter': { 'title': 'PNG Converter', 'description': 'Convert any image format to PNG.', 'icon': 'file-image', 'accept': '.jpg,.jpeg,.bmp,.webp,.tiff', 'allowed_extensions': ['.jpg', '.jpeg', '.bmp', '.webp', '.tiff'], 'category': 'image-conv', 'color': '#276749', 'gradient': 'from-green-500 to-emerald-700', 'target': 'png' },
    'bmp-converter': { 'title': 'BMP Converter', 'description': 'Convert any image format to Windows Bitmap.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png', '.webp'], 'category': 'image-conv', 'color': '#c05621', 'gradient': 'from-orange-500 to-red-500', 'target': 'bmp' },
    'gif-converter': { 'title': 'GIF Converter', 'description': 'Convert static images to GIF format.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#6b46c1', 'gradient': 'from-purple-500 to-indigo-700', 'target': 'gif' },
    'pdf-converter': { 'title': 'PDF Converter', 'description': 'Convert your images directly into a PDF document.', 'icon': 'file-text', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#dc2626', 'gradient': 'from-red-500 to-red-700', 'target': 'pdf' },
    'tiff-converter': { 'title': 'TIFF Converter', 'description': 'High-quality TIFF conversion for professional printing.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#0d9488', 'gradient': 'from-teal-500 to-teal-700', 'target': 'tiff' },
    'webp-converter': { 'title': 'WEBP Converter', 'description': 'Optimize your images for the web with WEBP format.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#0ea5e9', 'gradient': 'from-sky-500 to-sky-700', 'target': 'webp' },
    'dng-converter': { 'title': 'DNG Converter', 'description': 'DNG Digital Negative conversion placeholder.', 'icon': 'file-image', 'accept': '.*', 'allowed_extensions': ['.jpg', '.jpeg', '.png'], 'category': 'image-conv', 'color': '#111827', 'gradient': 'from-gray-700 to-black', 'target': 'tiff' },
}

def tool_page(request, tool_slug):

    if tool_slug not in IMAGE_TOOLS:
        raise Http404("Tool not found")

    tool = IMAGE_TOOLS[tool_slug]

    context = {
        'tool': tool,
        'tool_slug': tool_slug,

        # Page heading
        's1': tool.get(
            's1',
            tool['title']
        ),

        'highlight': tool.get(
            'highlight',
            'Online'
        ),

        # SEO
        'seo_title': tool.get(
            'seo_title',
            tool['title']
        ),

        'seo_description': tool.get(
            'seo_description',
            tool['description']
        ),

        'seo_keywords': tool.get(
            'seo_keywords',
            ''
        ),

        'seo_h1': tool.get(
            'seo_h1',
            f"{tool['title']} Online"
        ),

        'seo_intro': tool.get(
            'seo_intro',
            tool['description']
        ),

        'page_title': tool.get(
            'seo_title',
            f'{tool["title"]} — Image Editor'
        ),
    }

    return render(
        request,
        'image_processor/tool_detail.html',
        context
    )

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

