import os
import uuid
import math
from pydub import AudioSegment
import imageio_ffmpeg
from django.conf import settings
from converter.utils import format_download_name, ensure_media_dirs

# Fix for WinError 2: Set pydub's dependencies using imageio_ffmpeg
AudioSegment.converter = imageio_ffmpeg.get_ffmpeg_exe()
# Also try to find ffprobe if available in the same package (pydub uses it for info)
# Some versions of imageio_ffmpeg don't provide ffprobe but we'll try something similar 
# or hope ffmpeg is enough for basic tasks. 
# pydub often works with just ffmpeg if it's set correctly.

def get_output_path(original_name, target_extension, prefix=''):
    upload_dir, output_dir = ensure_media_dirs()
    base_name = os.path.splitext(original_name)[0]
    unique_suffix = uuid.uuid4().hex[:8].upper()
    filename = f"{base_name}{prefix}_{unique_suffix}.{target_extension}"
    return os.path.join(output_dir, filename)

def process_audio(input_path, original_name, tool_params):
    """
    Apply multiple audio effects in a pipeline.
    """
    audio = AudioSegment.from_file(input_path)
    
    # 1. Trimming
    start_time = tool_params.get('start')
    end_time = tool_params.get('end')
    if start_time is not None or end_time is not None:
        try:
            start_ms = float(start_time or 0) * 1000
            total_ms = len(audio)
            end_ms = float(end_time) * 1000 if end_time else total_ms
            if end_ms > total_ms: end_ms = total_ms
            if start_ms < 0: start_ms = 0
            if end_ms > start_ms:
                audio = audio[start_ms:end_ms]
        except (ValueError, TypeError):
            pass

    # 2. Volume
    try:
        volume_level = float(tool_params.get('volume', 100)) # percentage
        if volume_level != 100:
            if volume_level > 0:
                gain_db = 20 * math.log10(volume_level / 100.0)
                audio = audio.apply_gain(gain_db)
            else:
                audio = audio - 120 # Mute
    except (ValueError, TypeError):
        pass

    # 3. Speed (and Pitch if simple)
    try:
        speed = float(tool_params.get('speed', 1.0))
        if speed != 1.0:
            new_sample_rate = int(audio.frame_rate * speed)
            audio = audio._spawn(audio.raw_data, overrides={'frame_rate': new_sample_rate})
            audio = audio.set_frame_rate(audio.frame_rate)
    except (ValueError, TypeError):
        pass

    # 4. Pitch
    try:
        pitch = float(tool_params.get('pitch', 0))
        if pitch != 0:
            # Shift pitch by changing sample rate
            new_sample_rate = int(audio.frame_rate * (2.0 ** (pitch / 12.0)))
            audio = audio._spawn(audio.raw_data, overrides={'frame_rate': new_sample_rate})
            audio = audio.set_frame_rate(audio.frame_rate)
    except (ValueError, TypeError):
        pass

    # 5. Equalizer (Presets)
    preset = tool_params.get('preset', 'none')
    if preset != 'none':
        if preset == 'bass-boost':
            audio = audio.low_pass_filter(250).apply_gain(6)
        elif preset == 'treble-boost':
            audio = audio.high_pass_filter(5000).apply_gain(6)
        elif preset == 'classic':
            audio = audio.high_pass_filter(1000).low_pass_filter(4000).apply_gain(3)
        elif preset == 'dance':
            audio = audio.low_pass_filter(200).apply_gain(4).high_pass_filter(6000).apply_gain(4)
        elif preset == 'pop':
            audio = audio.high_pass_filter(1000).apply_gain(2)
        elif preset == 'rock':
            audio = audio.low_pass_filter(150).apply_gain(3).high_pass_filter(3000).apply_gain(3)

    # 5. Reverse
    if tool_params.get('reverse') == 'true':
        audio = audio.reverse()

    target_format = tool_params.get('format', 'mp3')
    output_path = get_output_path(original_name, target_format, '_processed')
    audio.export(output_path, format=target_format)
    return output_path

def merge_audios(input_paths, original_name, target_format='mp3'):
    combined = AudioSegment.empty()
    for path in input_paths:
        track = AudioSegment.from_file(path)
        combined += track
    
    output_path = get_output_path(original_name, target_format, '_merged')
    combined.export(output_path, format=target_format)
    return output_path
