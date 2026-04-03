import os
from moviepy.editor import VideoFileClip, AudioFileClip, concatenate_videoclips, vfx
import uuid
import time
from django.conf import settings
from converter.utils import format_download_name

def ensure_temp_dir():
    temp_dir = os.path.join(settings.MEDIA_ROOT, 'temp_video')
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir

def get_output_path(original_name, target_extension, prefix=''):
    base_name = os.path.splitext(original_name)[0]
    unique_suffix = uuid.uuid4().hex[:8].upper()
    filename = f"{base_name}{prefix}_{unique_suffix}.{target_extension}"
    return os.path.join(ensure_temp_dir(), filename)

def process_video(input_path, original_name, tool_params):
    """
    Main dispatcher for video tasks using MoviePy.
    """
    clip = VideoFileClip(input_path)
    tool = tool_params.get('tool')
    
    if tool == 'trim-video':
        start_s = float(tool_params.get('start', 0))
        end_s = float(tool_params.get('end', clip.duration))
        clip = clip.subclip(start_s, end_s)

    elif tool == 'crop-video':
        # Area in x1, y1, x2, y2 or center-based
        x = float(tool_params.get('x', 0))
        y = float(tool_params.get('y', 0))
        w = float(tool_params.get('w', clip.w))
        h = float(tool_params.get('h', clip.h))
        # moviepy crop: crop(x1, y1, x2, y2)
        clip = clip.crop(x1=x, y1=y, width=w, height=h)

    elif tool == 'rotate-video':
        angle = int(tool_params.get('angle', 90))
        clip = clip.rotate(angle)

    elif tool == 'change-speed':
        speed = float(tool_params.get('speed', 1.0))
        if speed != 1.0:
            clip = clip.fx(vfx.speedx, speed)

    elif tool == 'change-volume':
        vol = float(tool_params.get('volume', 1.0))
        clip = clip.volumex(vol)

    elif tool == 'reverse-video':
        clip = clip.fx(vfx.time_mirror)

    elif tool == 'loop-video':
        times = int(tool_params.get('loops', 2))
        clip = clip.fx(vfx.loop, n=times)

    target_format = tool_params.get('format', 'mp4')
    output_path = get_output_path(original_name, target_format, f'_{tool}')
    
    # Use ultrafast preset for interactive speed
    write_args = {'codec': 'libx264', 'audio_codec': 'aac', 'preset': 'ultrafast'}
    if target_format == 'webm':
        write_args = {'codec': 'libvpx', 'audio_codec': 'libvorbis', 'preset': 'ultrafast'}

    clip.write_videofile(output_path, **write_args)
    
    # Cleanup logic
    clip.close()
    time.sleep(0.5) # Windows handle release
    
    return output_path

def merge_videos(input_paths, original_name, target_format='mp4'):
    clips = [VideoFileClip(p) for p in input_paths]
    final_clip = concatenate_videoclips(clips, method="compose")
    
    output_path = get_output_path(original_name, target_format, '_merged')
    final_clip.write_videofile(output_path, codec='libx264', audio_codec='aac', preset='ultrafast')
    
    # Cleanup
    for c in clips: c.close()
    final_clip.close()
    time.sleep(0.5)
    
    return output_path
