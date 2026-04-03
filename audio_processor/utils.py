import os
from pydub import AudioSegment
import uuid
from django.conf import settings
from converter.utils import format_download_name, ensure_media_dirs


def get_output_path(original_name, target_extension, prefix=''):
    upload_dir, output_dir = ensure_media_dirs()
    base_name = os.path.splitext(original_name)[0]
    unique_suffix = uuid.uuid4().hex[:8].upper()
    filename = f"{base_name}{prefix}_{unique_suffix}.{target_extension}"
    return os.path.join(output_dir, filename)

def process_audio(input_path, original_name, tool_params):
    """
    Main dispatcher for audio tasks.
    """
    tool = tool_params.get('tool')
    
    if tool == 'video-to-audio':
        from moviepy.editor import VideoFileClip
        clip = VideoFileClip(input_path)
        output_path = get_output_path(original_name, 'mp3', '_extracted')
        clip.audio.write_audiofile(output_path)
        clip.close()
        return output_path

    # All other tools use pydub
    audio = AudioSegment.from_file(input_path)
    
    if tool == 'trim-audio':
        start_ms = float(tool_params.get('start', 0)) * 1000
        end_ms = float(tool_params.get('end', audio.duration_seconds)) * 1000
        audio = audio[start_ms:end_ms]
        
        # Fades
        fade_in = float(tool_params.get('fade_in', 0)) * 1000
        fade_out = float(tool_params.get('fade_out', 0)) * 1000
        if fade_in > 0: audio = audio.fade_in(int(fade_in))
        if fade_out > 0: audio = audio.fade_out(int(fade_out))

    elif tool == 'change-volume':
        volume_change = float(tool_params.get('volume', 100)) - 100 # percentage
        # pydub volume is in dB. 6dB is roughly double volume.
        # Simple linear to dB conversion helper
        if volume_change != 0:
            audio = audio + (volume_change / 10) # rough mapping

    elif tool == 'change-speed':
        speed = float(tool_params.get('speed', 1.0))
        if speed != 1.0:
            # Note: This changes PITCH too. For pitch-invariant speed, we need rubberband or librosa.
            new_sample_rate = int(audio.frame_rate * speed)
            audio = audio._spawn(audio.raw_data, overrides={'frame_rate': new_sample_rate})
            audio = audio.set_frame_rate(audio.frame_rate)

    elif tool == 'reverse-audio':
        audio = audio.reverse()

    elif tool == 'audio-equalizer':
        preset = tool_params.get('preset', 'none')
        if preset == 'bass-boost':
            audio = audio.low_pass_filter(250).apply_gain(6) + audio
        elif preset == 'treble-boost':
            audio = audio.high_pass_filter(5000).apply_gain(6) + audio


    target_format = tool_params.get('format', 'mp3')
    output_path = get_output_path(original_name, target_format, f'_{tool}')
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
