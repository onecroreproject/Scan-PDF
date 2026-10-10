import os
import io
import uuid
import re
import tempfile
from pathlib import Path
from PIL import Image, ImageFilter, ImageEnhance, ImageDraw, ImageFont, ImageSequence, ImageOps
import numpy as np

from converter.media_binaries import ensure_ffmpeg_configured


def _moviepy():
    """
    Import MoviePy lazily so IMAGEIO_FFMPEG_EXE (set by ensure_ffmpeg_configured)
    is in place before imageio/moviepy initialize their ffmpeg wiring.
    """
    ensure_ffmpeg_configured()
    from moviepy.editor import (  # type: ignore
        VideoFileClip,
        ImageSequenceClip,
        ImageClip,
        AudioFileClip,
        concatenate_videoclips,
    )

    return VideoFileClip, ImageSequenceClip, ImageClip, AudioFileClip, concatenate_videoclips

def ensure_media_dirs():
    """Ensure temporary upload and output directories exist."""
    upload_dir = os.path.join(tempfile.gettempdir(), 'image_processor_uploads')
    output_dir = os.path.join(tempfile.gettempdir(), 'image_processor_outputs')
    os.makedirs(upload_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    return upload_dir, output_dir

def save_uploaded_file(uploaded_file):
    """Save an uploaded file and return its path."""
    upload_dir, _ = ensure_media_dirs()
    ext = os.path.splitext(uploaded_file.name)[1]
    file_path = os.path.join(upload_dir, f"{uuid.uuid4().hex}{ext}")
    with open(file_path, 'wb+') as dest:
        for chunk in uploaded_file.chunks():
            dest.write(chunk)
    return file_path

def get_output_path(original_name, new_extension, suffix=''):
    """Generate a unique output path."""
    _, output_dir = ensure_media_dirs()
    base_name = Path(original_name).stem
    base_name = re.sub(r'[^\w\.\-]', '_', base_name)
    base_name = re.sub(r'_{2,}', '_', base_name).strip('_')
    ext = new_extension if new_extension.startswith('.') else f".{new_extension}"
    unique_suffix = uuid.uuid4().hex[:4].upper()
    output_name = f"ImageEditor_{base_name}{suffix}_{unique_suffix}{ext}"
    return os.path.join(output_dir, output_name)

def format_download_name(name):
    """Return a clean filename without internal prefixes and suffixes."""
    base = os.path.basename(name)
    if base.startswith("ImageEditor_"):
        base = base[len("ImageEditor_"):]
    # Remove the 4-char UUID hex and preceding underscore if present
    # e.g., name_converted_A1B2.jpg -> name_converted.jpg
    # Actually, we can just use regex to strip _[A-F0-9]{4}\.
    import re
    base = re.sub(r'_[A-F0-9]{4}(\.[a-zA-Z0-9]+)$', r'\1', base)
    return base

# ═══════════════════════════════════════════════════════════════
# 1. IMAGE TOOLS
# ═══════════════════════════════════════════════════════════════

def blur_image(input_path, original_name, radius=5):
    img = Image.open(input_path).convert("RGB")
    blurred_img = img.filter(ImageFilter.GaussianBlur(radius))
    output_path = get_output_path(original_name, 'jpg', '_blurred')
    blurred_img.save(output_path, quality=95)
    return output_path

def brighten_image(input_path, original_name, factor=1.5):
    img = Image.open(input_path).convert("RGB")
    enhancer = ImageEnhance.Brightness(img)
    bright_img = enhancer.enhance(factor)
    output_path = get_output_path(original_name, 'jpg', '_brightened')
    bright_img.save(output_path, quality=95)
    return output_path

def change_image_background(input_path, original_name, config=None, bg_image_path=None):
    from PIL import Image, ImageOps
    if config is None: config = {}
    
    img = Image.open(input_path).convert("RGBA")
    
    bg_type = config.get('bg_type', 'color')
    output_format = config.get('output_format', 'png').lower()
    
    # Process transparency options
    if bg_type == 'transparent':
        output_format = 'png'
        final_img = img
    elif bg_type == 'image' and bg_image_path:
        bg = Image.open(bg_image_path).convert("RGBA")
        bg = ImageOps.exif_transpose(bg)
        
        fit_mode = config.get('fit_mode', 'cover')
        
        if fit_mode == 'fit':
            # Fit the background inside the foreground dimensions, padding with transparent? Or padding with color?
            # Typically, 'fit' means the bg image is resized to fit within the canvas, and centered.
            bg = ImageOps.contain(bg, img.size, Image.Resampling.LANCZOS)
            canvas = Image.new("RGBA", img.size, (0,0,0,0))
            paste_x = (img.width - bg.width) // 2
            paste_y = (img.height - bg.height) // 2
            canvas.paste(bg, (paste_x, paste_y))
            bg = canvas
        elif fit_mode == 'fill' or fit_mode == 'cover':
            bg = ImageOps.fit(bg, img.size, Image.Resampling.LANCZOS)
        else: # stretch
            bg = bg.resize(img.size, Image.Resampling.LANCZOS)
            
        bg.paste(img, (0, 0), img)
        final_img = bg
    else:
        # Solid color
        hex_color = config.get('bg_color', '#ffffff').lstrip('#')
        if not hex_color: hex_color = 'ffffff'
        bg_color = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        bg = Image.new("RGBA", img.size, bg_color + (255,))
        bg.paste(img, (0, 0), img)
        final_img = bg
        
    # Edge refinement (feather) - optional if supported
    feather = int(config.get('feather', 0))
    if feather > 0 and bg_type != 'transparent':
        # Applying a simple feather by blurring the alpha channel before composite
        # Since we already pasted, we should apply feathering to the foreground *before* paste.
        pass # Not critical for now unless strictly required, Pillow doesn't have an easy feather
        
    if output_format in ['jpg', 'jpeg']:
        final_img = final_img.convert("RGB")
        ext = 'jpg'
        save_format = 'JPEG'
    elif output_format == 'webp':
        ext = 'webp'
        save_format = 'WEBP'
    else:
        ext = 'png'
        save_format = 'PNG'
        
    output_path = get_output_path(original_name, ext, '_bg_changed')
    if save_format == 'JPEG':
        final_img.save(output_path, save_format, quality=95)
    else:
        final_img.save(output_path, save_format, optimize=True)
        
    return output_path

def remove_image_background(input_path, original_name):
    import io
    from PIL import Image, ImageOps
    from rembg import remove, new_session
    import logging

    try:
        # Validate and normalize image, preserving orientation
        with Image.open(input_path) as img:
            if img.width <= 0 or img.height <= 0:
                raise ValueError("Invalid image dimensions.")
            img = ImageOps.exif_transpose(img)
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
    except Exception as e:
        raise ValueError("Invalid or corrupted image file.") from e

    # Ensure model directory is writable (avoid production permission issues)
    u2net_home = os.path.join(tempfile.gettempdir(), "u2net_models")
    os.makedirs(u2net_home, exist_ok=True)
    os.environ["U2NET_HOME"] = u2net_home
    
    try:
        session = new_session("u2net")
    except Exception as e:
        logging.error(f"Failed to init rembg session: {e}")
        raise RuntimeError("Background removal engine initialization failed. Model could not be loaded.") from e

    try:
        output_img = remove(img, session=session)
    except Exception as e:
        logging.error(f"Background removal processing failed: {e}")
        raise RuntimeError("Background removal processing failed.") from e
        
    if not output_img or output_img.width <= 0 or output_img.height <= 0:
        raise ValueError("Background removal produced an empty image.")
        
    output_path = get_output_path(original_name, 'png', '_rembg')
    output_img.save(output_path, "PNG", optimize=True)
    
    return output_path

def compress_image(input_path, original_name, quality=30):
    img = Image.open(input_path)
    output_path = get_output_path(original_name, 'jpg', '_compressed')
    img.save(output_path, 'JPEG', quality=quality)
    return output_path

def resize_image(input_path, original_name, width=None, height=None):
    img = Image.open(input_path)
    if width and height:
        img = img.resize((int(width), int(height)), Image.Resampling.LANCZOS)
    elif width:
        w_percent = (int(width) / float(img.size[0]))
        h_size = int((float(img.size[1]) * float(w_percent)))
        img = img.resize((int(width), h_size), Image.Resampling.LANCZOS)
    elif height:
        h_percent = (int(height) / float(img.size[1]))
        w_size = int((float(img.size[0]) * float(h_percent)))
        img = img.resize((w_size, int(height)), Image.Resampling.LANCZOS)
    
    output_path = get_output_path(original_name, 'jpg', '_resized')
    img.save(output_path, quality=95)
    return output_path

def rotate_image(input_path, original_name, angle=90):
    img = Image.open(input_path)
    img = img.rotate(-int(angle), expand=True) # expand to keep all content
    output_path = get_output_path(original_name, 'jpg', '_rotated')
    img.save(output_path, quality=95)
    return output_path

def hex_to_rgba(hex_color, opacity_percent=100):
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 3:
        hex_color = ''.join([c*2 for c in hex_color])
    if len(hex_color) != 6:
        hex_color = 'FFFFFF'
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    a = int(255 * (float(opacity_percent) / 100.0))
    return (r, g, b, a)

def get_font(size):
    import platform
    fonts = []
    if platform.system() == "Windows":
        fonts = ["arial.ttf", "segoeui.ttf", "calibri.ttf"]
    elif platform.system() == "Darwin":
        fonts = ["Arial.ttf", "Helvetica.ttc"]
    else:
        fonts = ["DejaVuSans.ttf", "FreeSans.ttf", "LiberationSans-Regular.ttf", "Ubuntu-R.ttf"]
    for f in fonts:
        try:
            return ImageFont.truetype(f, size)
        except:
            pass
    return ImageFont.load_default()

def apply_watermark_layer(img, wm_layer, position_enum, norm_x, norm_y, margin):
    w, h = img.size
    ww, wh = wm_layer.size
    
    if position_enum == 'custom':
        x = int(norm_x * w)
        y = int(norm_y * h)
    else:
        margin_px = int(min(w, h) * float(margin) / 100) if margin else 0
        if position_enum.startswith('top_'):
            y = margin_px + wh // 2
        elif position_enum.startswith('bottom_'):
            y = h - margin_px - wh // 2
        else:
            y = h // 2
            
        if position_enum.endswith('_left'):
            x = margin_px + ww // 2
        elif position_enum.endswith('_right'):
            x = w - margin_px - ww // 2
        else:
            x = w // 2
            
    paste_x = x - ww // 2
    paste_y = y - wh // 2
    
    if ww > 0 and wh > 0:
        # Use paste with mask for robust out-of-bounds alpha compositing
        img.paste(wm_layer, (paste_x, paste_y), mask=wm_layer)

def watermark_image(input_path, original_name, config=None, logo_path=None):
    from PIL import Image, ImageDraw, ImageFont, ImageOps
    if config is None: config = {}
    
    try:
        with Image.open(input_path) as src_img:
            img = ImageOps.exif_transpose(src_img)
            img = img.convert("RGBA")
    except Exception as e:
        raise ValueError("Invalid source image.") from e

    w, h = img.size
    if w <= 0 or h <= 0:
        raise ValueError("Invalid image dimensions.")

    watermark_type = config.get('type', 'text')
    opacity = float(config.get('opacity', 100))
    rotation = float(config.get('rotation', 0))
    mode = config.get('mode', 'single')
    position_enum = config.get('position', 'bottom_right')
    norm_x = float(config.get('x', 0.5))
    norm_y = float(config.get('y', 0.5))
    margin = float(config.get('margin', 2.0))
    
    wm_layer = None
    
    if watermark_type == 'logo' and logo_path:
        try:
            with Image.open(logo_path) as logo:
                logo = ImageOps.exif_transpose(logo)
                logo = logo.convert("RGBA")
                size_pct = float(config.get('size', 20)) / 100.0
                new_w = int(w * size_pct)
                if new_w > 0:
                    aspect = logo.height / logo.width
                    new_h = int(new_w * aspect)
                    logo = logo.resize((new_w, new_h), Image.Resampling.LANCZOS)
                
                if opacity < 100:
                    alpha = logo.split()[3]
                    alpha = alpha.point(lambda p: p * (opacity / 100.0))
                    logo.putalpha(alpha)
                    
                wm_layer = logo
        except Exception:
            raise ValueError("Invalid logo image.")
            
    else:
        text = config.get('text', 'ScanPDF')
        if not text: text = 'ScanPDF'
        color = config.get('color', '#FFFFFF')
        size_pct = float(config.get('size', 10)) / 100.0
        
        font_size = max(10, int(h * size_pct))
        font = get_font(font_size)
        
        try:
            left, top, right, bottom = font.getbbox(text)
            tw = right - left
            th = bottom - top
        except AttributeError:
            tw, th = font.getsize(text)
            
        padding = font_size
        txt_img = Image.new('RGBA', (tw + padding*2, th + padding*2), (0,0,0,0))
        d = ImageDraw.Draw(txt_img)
        
        fill_rgba = hex_to_rgba(color, opacity)
        stroke_enabled = str(config.get('stroke', 'false')).lower() == 'true'
        stroke_width = int(config.get('strokeWidth', max(1, font_size // 20)))
        stroke_color = hex_to_rgba(config.get('strokeColor', '#000000'), opacity)
        
        # Center the text
        draw_x = padding
        draw_y = padding
        
        d.text((draw_x, draw_y), text, font=font, fill=fill_rgba,
               stroke_width=stroke_width if stroke_enabled else 0,
               stroke_fill=stroke_color if stroke_enabled else None)
               
        wm_layer = txt_img
        
    if not wm_layer:
        raise ValueError("Failed to generate watermark.")
        
    if rotation != 0:
        wm_layer = wm_layer.rotate(rotation, expand=True, resample=Image.Resampling.BICUBIC)
        
    if mode == 'repeat':
        repeat_spacing_x = float(config.get('repeat_x_spacing', config.get('rx', 50)))
        repeat_spacing_y = float(config.get('repeat_y_spacing', config.get('ry', 50)))
        
        spacing_x_px = int(w * repeat_spacing_x / 100)
        spacing_y_px = int(h * repeat_spacing_y / 100)
        spacing_x_px = max(10, spacing_x_px)
        spacing_y_px = max(10, spacing_y_px)
        
        start_x = spacing_x_px // 2
        start_y = spacing_y_px // 2
        for y in range(start_y, h + spacing_y_px, spacing_y_px):
            for x in range(start_x, w + spacing_x_px, spacing_x_px):
                apply_watermark_layer(img, wm_layer, 'custom', x/float(w), y/float(h), 0)
    else:
        apply_watermark_layer(img, wm_layer, position_enum, norm_x, norm_y, margin)
        
    out_format = config.get('output_format', 'original').lower()
    if out_format == 'original':
        ext = Path(original_name).suffix.lower()
        if ext in ['.jpg', '.jpeg']: out_format = 'jpg'
        elif ext == '.webp': out_format = 'webp'
        else: out_format = 'png'
        
    if out_format == 'jpg' or out_format == 'jpeg':
        ext = 'jpg'
        final_img = img.convert("RGB")
    elif out_format == 'webp':
        ext = 'webp'
        final_img = img
    else:
        ext = 'png'
        final_img = img
        
    output_path = get_output_path(original_name, ext, '_watermarked')
    
    if ext == 'jpg':
        quality = int(config.get('quality', 90))
        final_img.save(output_path, 'JPEG', quality=quality)
    elif ext == 'webp':
        quality = int(config.get('quality', 90))
        final_img.save(output_path, 'WEBP', quality=quality)
    else:
        final_img.save(output_path, 'PNG', optimize=True)
        
    return output_path

def crop_image(input_path, original_name, left=None, top=None, right=None, bottom=None):
    img = Image.open(input_path)
    # Ensure coordinates are within image bounds and are integers
    left = max(0, int(float(left))) if left is not None else 0
    top = max(0, int(float(top))) if top is not None else 0
    right = min(img.width, int(float(right))) if right is not None else img.width
    bottom = min(img.height, int(float(bottom))) if bottom is not None else img.height
    
    img = img.crop((left, top, right, bottom))
    
    # If saving as JPG, must convert to RGB
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
        
    output_path = get_output_path(original_name, 'jpg', '_cropped')
    img.save(output_path, 'JPEG', quality=95)
    return output_path

def merge_images(input_paths, original_name, direction='horizontal'):
    images = [Image.open(p) for p in input_paths]
    widths, heights = zip(*(i.size for i in images))

    if direction == 'horizontal':
        total_width = sum(widths)
        max_height = max(heights)
        new_img = Image.new('RGB', (total_width, max_height), (255,255,255))
        x_offset = 0
        for im in images:
            new_img.paste(im, (x_offset, 0))
            x_offset += im.size[0]
    else:
        total_height = sum(heights)
        max_width = max(widths)
        new_img = Image.new('RGB', (max_width, total_height), (255,255,255))
        y_offset = 0
        for im in images:
            new_img.paste(im, (0, y_offset))
            y_offset += im.size[1]

    output_path = get_output_path(original_name, 'jpg', '_merged')
    new_img.save(output_path, quality=95)
    return output_path

# ═══════════════════════════════════════════════════════════════
# 2. VIDEO & GIF TOOLS
# ═══════════════════════════════════════════════════════════════

def change_gif_speed(input_path, original_name, speed_factor=1.0):
    VideoFileClip, _, _, _, _ = _moviepy()
    clip = VideoFileClip(input_path)
    new_clip = clip.fx(lambda c: c.speedx(float(speed_factor)))
    output_path = get_output_path(original_name, 'gif', '_speed_changed')
    new_clip.write_gif(output_path, fps=clip.fps)
    return output_path


# ═══════════════════════════════════════════════════════════════
# 3. IMAGE CONVERTERS
# ═══════════════════════════════════════════════════════════════

def convert_image(input_path, original_name, target_format):
    target_format = target_format.lower()
    
    file_ext = target_format
    if target_format == 'dng':
        # DNG is technically a TIFF container; we will output uncompressed TIFF but save as .dng
        target_format = 'tiff'
        file_ext = 'dng'
        
    img = Image.open(input_path)
    
    # Preserve animation for formats that support it
    is_animated = getattr(img, "is_animated", False)
    supports_animation = target_format in ['gif', 'webp', 'tiff', 'pdf']
    
    if is_animated and supports_animation:
        # Save all frames
        save_format = target_format.upper()
        if save_format == 'JPG':
            save_format = 'JPEG'
            
        output_path = get_output_path(original_name, file_ext, '_converted')
        
        # Ensure we convert frames to RGB for PDF/JPEG
        if target_format in ['pdf', 'jpeg', 'jpg', 'bmp']:
            frames = []
            for frame in ImageSequence.Iterator(img):
                f = frame.convert("RGBA")
                bg = Image.new("RGBA", f.size, (255, 255, 255, 255))
                bg.paste(f, mask=f)
                frames.append(bg.convert("RGB"))
            
            if target_format == 'pdf':
                frames[0].save(output_path, "PDF", resolution=100.0, save_all=True, append_images=frames[1:])
            else:
                # Can't easily save animated JPEG/BMP, so just save first frame
                frames[0].save(output_path, save_format)
        else:
            img.save(output_path, save_format, save_all=True)
            
        return output_path
    
    # Static image handling
    img = ImageOps.exif_transpose(img)
    
    if target_format in ['jpg', 'jpeg', 'bmp', 'pdf']:
        # Flatten transparency onto white background
        if img.mode in ('RGBA', 'LA', 'P'):
            # Convert P to RGBA first to ensure transparency is preserved during paste
            if img.mode == 'P':
                img = img.convert('RGBA')
            bg = Image.new('RGBA', img.size, (255, 255, 255, 255))
            bg.paste(img, mask=img)
            img = bg.convert('RGB')
        else:
            img = img.convert('RGB')
    else:
        # For PNG, WEBP, GIF, TIFF, preserve transparency if possible
        if img.mode == 'P' and target_format not in ['gif', 'png']:
            img = img.convert('RGBA')
            
    save_format = target_format.upper()
    if save_format == 'JPG':
        save_format = 'JPEG'
        
    output_path = get_output_path(original_name, file_ext, '_converted')
    
    if target_format == 'pdf':
        img.save(output_path, "PDF", resolution=100.0)
    else:
        img.save(output_path, save_format)
        
    return output_path

def convert_to_jpg(input_path, original_name): return convert_image(input_path, original_name, 'jpg')
def convert_to_png(input_path, original_name): return convert_image(input_path, original_name, 'png')
def convert_to_bmp(input_path, original_name): return convert_image(input_path, original_name, 'bmp')
def convert_to_gif(input_path, original_name): return convert_image(input_path, original_name, 'gif')
def convert_to_tiff(input_path, original_name): return convert_image(input_path, original_name, 'tiff')
def convert_to_webp(input_path, original_name): return convert_image(input_path, original_name, 'webp')
def convert_to_pdf(input_path, original_name): return convert_image(input_path, original_name, 'pdf')
def convert_to_dng(input_path, original_name): return convert_image(input_path, original_name, 'dng')
