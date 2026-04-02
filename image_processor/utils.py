import os
import io
import uuid
import re
import tempfile
from pathlib import Path
from PIL import Image, ImageFilter, ImageEnhance, ImageDraw, ImageFont
import numpy as np
import cv2

# Functionality for video and gif processing
from moviepy.editor import VideoFileClip, ImageSequenceClip

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
    """Clean filename for download."""
    stem = Path(name).stem
    ext = Path(name).suffix
    stem = re.sub(r'_[0-9a-fA-F]{4,32}$', '', stem)
    if not stem.lower().startswith('imageeditor'):
        stem = f"ImageEditor_{stem}"
    stem = re.sub(r'[^\w\.\-]', '_', stem)
    stem = re.sub(r'_{2,}', '_', stem).strip('_')
    return f"{stem}{ext}"

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

def change_image_background(input_path, original_name, bg_color=(255, 255, 255)):
    from rembg import remove
    with open(input_path, 'rb') as i:
        input_data = i.read()
    output_data = remove(input_data)
    img = Image.open(io.BytesIO(output_data)).convert("RGBA")
    
    # Create new background
    new_bg = Image.new("RGB", img.size, bg_color)
    # Paste using the alpha channel as a mask
    new_bg.paste(img, (0, 0), img)
    
    output_path = get_output_path(original_name, 'jpg', '_bg_changed')
    new_bg.save(output_path, 'JPEG', quality=95)
    return output_path

def remove_image_background(input_path, original_name):
    from rembg import remove
    with open(input_path, 'rb') as i:
        input_data = i.read()
    output_data = remove(input_data)
    output_path = get_output_path(original_name, 'png', '_rembg')
    with open(output_path, 'wb') as f:
        f.write(output_data)
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

def watermark_image(input_path, original_name, text="ScanPDF", opacity=128):
    img = Image.open(input_path).convert("RGBA")
    txt = Image.new('RGBA', img.size, (255, 255, 255, 0))
    
    # Attempt to find a font
    try:
        font = ImageFont.truetype("arial.ttf", int(img.size[0] / 10))
    except:
        font = ImageFont.load_default()
        
    d = ImageDraw.Draw(txt)
    # Position in center
    w, h = img.size
    d.text((w/2, h/2), text, fill=(255, 255, 255, opacity), font=font, anchor="mm")
    
    combined = Image.alpha_composite(img, txt)
    output_path = get_output_path(original_name, 'jpg', '_watermarked')
    combined.convert("RGB").save(output_path, quality=95)
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
    clip = VideoFileClip(input_path)
    new_clip = clip.fx(lambda c: c.speedx(float(speed_factor)))
    output_path = get_output_path(original_name, 'gif', '_speed_changed')
    new_clip.write_gif(output_path, fps=clip.fps)
    return output_path

def extract_image_from_video(input_path, original_name, timestamp=1.0):
    clip = VideoFileClip(input_path)
    frame = clip.get_frame(float(timestamp))
    output_path = get_output_path(original_name, 'jpg', f'_frame_{timestamp}')
    img = Image.fromarray(frame)
    img.save(output_path, quality=95)
    return output_path

def gif_to_video(input_path, original_name, target_format='mp4', duration='default', effect='none', effect_duration='3', speed_factor='default', bg_color='default', music_path=None):
    from moviepy.editor import VideoFileClip, AudioFileClip, CompositeVideoClip, ColorClip, vfx
    import moviepy.video.fx.all as vfx
    
    clip = VideoFileClip(input_path)
    
    # 1. Handle Speed
    if speed_factor != 'default':
        clip = clip.fx(vfx.speedx, float(speed_factor))
    
    # 2. Handle Duration
    if duration != 'default':
        target_dur = float(duration)
        if clip.duration < target_dur:
            clip = clip.fx(vfx.loop, duration=target_dur)
        else:
            clip = clip.subclip(0, target_dur)
    
    w, h = clip.size
    fps = clip.fps or 24
    
    # 3. Handle Color Background / Theme
    colors = {
        'azure': (240, 255, 255), 'black': (0,0,0), 'blue': (0,0,255), 'brown': (165,42,42),
        'cyan': (0,255,255), 'fuchsia': (255,0,255), 'gold': (255,215,0), 'gray': (128,128,128),
        'green': (0,128,0), 'maroon': (128,0,0), 'navy': (0,0,128), 'olive': (128,128,0),
        'orange': (255,166,0), 'pink': (255,192,203), 'purple': (128,0,128), 'red': (255,0,0),
        'silver': (192,192,192), 'skyblue': (135,206,235), 'white': (255,255,255)
    }
    
    if bg_color.lower() in colors:
        bg_rgb = colors[bg_color.lower()]
        bg_clip = ColorClip(size=(w, h), color=bg_rgb, duration=clip.duration)
        clip = CompositeVideoClip([bg_clip, clip.set_position('center')])

    # 4. Handle Effects
    eff_dur = float(effect_duration)
    if clip.duration < eff_dur: eff_dur = clip.duration

    if effect != 'none':
        if effect.startswith('zoom_in'):
            clip = clip.fx(vfx.resize, lambda t: 1 + 0.3 * (t/clip.duration))
        elif effect.startswith('zoom_out'):
            clip = clip.fx(vfx.resize, lambda t: 1.3 - 0.3 * (t/clip.duration))
        elif effect == 'pan_left':
            clip = clip.set_position(lambda t: (max(0, w - w*(1 + 0.1*t)), 'center'))
        elif effect == 'pan_right':
            clip = clip.set_position(lambda t: (min(0, -w*0.1*t), 'center'))
        elif effect == 'blur_to_clear':
            clip = clip.fx(vfx.gaussian_blur, lambda t: max(0, 10 * (1 - t/eff_dur)))
        elif effect == 'clear_to_blur':
            clip = clip.fx(vfx.gaussian_blur, lambda t: 10 * (t/eff_dur))
        elif effect == 'fade':
             clip = clip.fx(vfx.fadein, eff_dur).fx(vfx.fadeout, eff_dur)
        elif effect.startswith('rotate'):
            angle = 45 if '45' in effect else 90
            if 'l' in effect: angle = -angle
            clip = clip.fx(vfx.rotate, lambda t: angle * (t/clip.duration))
        elif 'slide' in effect or 'wipe' in effect:
            # Simplified transition for single clip
            clip = clip.fx(vfx.fadein, eff_dur)
        elif 'slice' in effect:
            clip = clip.fx(vfx.mirror_x) if 'left' in effect else clip

    # 5. Handle Music
    if music_path and os.path.exists(music_path):
        try:
            audio = AudioFileClip(music_path)
            if audio.duration < clip.duration:
                audio = audio.fx(vfx.loop, duration=clip.duration)
            else:
                audio = audio.subclip(0, clip.duration)
            clip = clip.set_audio(audio)
        except:
            pass

    output_path = get_output_path(original_name, target_format, f'_processed')
    
    write_args = {'codec': 'libx264', 'audio_codec': 'aac'}
    if target_format == 'webm':
        write_args = {'codec': 'libvpx', 'audio_codec': 'libvorbis'}
    
    clip.write_videofile(output_path, **write_args)
    return output_path

def image_to_video(input_paths, original_name, target_format='mp4', duration_per_image='2', transition_type='fade', music_path=None):
    from moviepy.editor import ImageClip, AudioFileClip, CompositeVideoClip, concatenate_videoclips
    
    clips = []
    dur = float(duration_per_image)
    
    # Process each image into a clip
    for path in input_paths:
        clip = ImageClip(path).set_duration(dur)
        if transition_type == 'fade':
            clip = clip.crossfadein(0.5).crossfadeout(0.5)
        clips.append(clip)
    
    # Concatenate with or without transitions
    if transition_type == 'fade':
        # Overlap clips for crossfade
        final_clip = concatenate_videoclips(clips, method="compose", padding=-0.5)
    elif 'slide' in transition_type:
        final_clip = concatenate_videoclips(clips, method="compose")
        # Basic slide can be implemented via position animation if needed, 
        # but concatenate method="compose" handles basic sequencing.
    else:
        final_clip = concatenate_videoclips(clips, method="chain")

    # Handle Music
    if music_path and os.path.exists(music_path):
        try:
            audio = AudioFileClip(music_path)
            if audio.duration < final_clip.duration:
                from moviepy.video.fx.all import loop
                audio = audio.fx(loop, duration=final_clip.duration)
            else:
                audio = audio.subclip(0, final_clip.duration)
            final_clip = final_clip.set_audio(audio)
        except:
            pass

    output_path = get_output_path(original_name, target_format, '_slideshow')
    
    write_args = {'codec': 'libx264', 'audio_codec': 'aac', 'fps': 24}
    if target_format == 'webm':
        write_args = {'codec': 'libvpx', 'audio_codec': 'libvorbis', 'fps': 24}
    
    final_clip.write_videofile(output_path, **write_args)
    return output_path

# ═══════════════════════════════════════════════════════════════
# 3. IMAGE CONVERTERS
# ═══════════════════════════════════════════════════════════════

def convert_image(input_path, original_name, target_format):
    img = Image.open(input_path).convert("RGB")
    target_format = target_format.lower()
    if target_format == 'pdf':
        output_path = get_output_path(original_name, 'pdf', '_converted')
        img.save(output_path, "PDF", resolution=100.0)
    else:
        save_format = target_format.upper()
        if save_format == 'JPG':
            save_format = 'JPEG'
        output_path = get_output_path(original_name, target_format, '_converted')
        img.save(output_path, save_format)
    return output_path

def convert_to_jpg(input_path, original_name): return convert_image(input_path, original_name, 'jpg')
def convert_to_png(input_path, original_name): return convert_image(input_path, original_name, 'png')
def convert_to_bmp(input_path, original_name): return convert_image(input_path, original_name, 'bmp')
def convert_to_gif(input_path, original_name): return convert_image(input_path, original_name, 'gif')
def convert_to_tiff(input_path, original_name): return convert_image(input_path, original_name, 'tiff')
def convert_to_webp(input_path, original_name): return convert_image(input_path, original_name, 'webp')
def convert_to_pdf(input_path, original_name): return convert_image(input_path, original_name, 'pdf')

# DNG (Digital Negative) is harder. Pillow doesn't write DNG natively. 
# It can read RAW if rawpy is installed. Let's stick to standard formats for now
# or use a placeholder if not possible safely.
def convert_to_dng(input_path, original_name):
    # DNG export is not trivial in Python without heavy libs like Wand/ImageMagick
    # For now, we'll convert to TIFF which is similar in some contexts or just raise error if called specifically.
    return convert_image(input_path, original_name, 'tiff')
