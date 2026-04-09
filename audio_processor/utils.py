import os
import uuid
import math
import shutil
import glob
from pydub import AudioSegment
from pydub.effects import speedup
import imageio_ffmpeg
from converter.utils import format_download_name, ensure_media_dirs

def _get_winget_ffmpeg_bin_dir():
    local_app_data = os.environ.get('LOCALAPPDATA', '')
    if not local_app_data:
        return None
    pattern = os.path.join(
        local_app_data,
        'Microsoft',
        'WinGet',
        'Packages',
        'Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe',
        'ffmpeg-*-full_build',
        'bin',
    )
    matches = sorted(glob.glob(pattern), reverse=True)
    return matches[0] if matches else None


def _configure_audio_binaries():
    """Configure ffmpeg/ffprobe paths for pydub, with Windows-safe fallbacks."""
    winget_bin_dir = _get_winget_ffmpeg_bin_dir()
    winget_ffmpeg = os.path.join(winget_bin_dir, 'ffmpeg.exe') if winget_bin_dir else None
    winget_ffprobe = os.path.join(winget_bin_dir, 'ffprobe.exe') if winget_bin_dir else None

    ffmpeg_candidates = []
    if winget_ffmpeg and os.path.exists(winget_ffmpeg):
        ffmpeg_candidates.append(winget_ffmpeg)
    which_ffmpeg = shutil.which('ffmpeg')
    if which_ffmpeg:
        ffmpeg_candidates.append(which_ffmpeg)
    try:
        ffmpeg_candidates.append(imageio_ffmpeg.get_ffmpeg_exe())
    except Exception:
        pass

    ffmpeg_path = next((p for p in ffmpeg_candidates if p and os.path.exists(p)), None)
    if not ffmpeg_path:
        raise RuntimeError(
            "FFmpeg binary is not available. Please install ffmpeg or keep imageio-ffmpeg installed."
        )

    ffprobe_candidates = []
    if winget_ffprobe and os.path.exists(winget_ffprobe):
        ffprobe_candidates.append(winget_ffprobe)
    which_ffprobe = shutil.which('ffprobe')
    if which_ffprobe:
        ffprobe_candidates.append(which_ffprobe)
    if not ffprobe_candidates:
        candidate = os.path.join(os.path.dirname(ffmpeg_path), 'ffprobe.exe')
        if os.path.exists(candidate):
            ffprobe_candidates.append(candidate)
    ffprobe_path = next((p for p in ffprobe_candidates if p and os.path.exists(p)), None)

    # Ensure ffmpeg/ffprobe directory is visible to pydub's `which()` checks.
    bin_dir = os.path.dirname(ffmpeg_path)
    current_path = os.environ.get('PATH', '')
    path_parts = current_path.split(os.pathsep) if current_path else []
    if bin_dir not in path_parts:
        os.environ['PATH'] = bin_dir + os.pathsep + current_path if current_path else bin_dir

    # pydub binary configuration
    AudioSegment.converter = ffmpeg_path
    AudioSegment.ffmpeg = ffmpeg_path
    if ffprobe_path:
        AudioSegment.ffprobe = ffprobe_path


def _load_audio_segment(input_path):
    """Load audio robustly and emit user-friendly binary errors."""
    ext = os.path.splitext(input_path)[1].lstrip('.').lower() or None
    try:
        return AudioSegment.from_file(input_path, format=ext)
    except FileNotFoundError as exc:
        # Self-heal path resolution in long-running server processes.
        try:
            _configure_audio_binaries()
            return AudioSegment.from_file(input_path, format=ext)
        except Exception:
            try:
                # Final fallback: let ffmpeg auto-detect format.
                return AudioSegment.from_file(input_path)
            except Exception:
                raise RuntimeError(
                    "Audio engine binary not found (ffmpeg/ffprobe). Restart the server once and try again."
                ) from exc
    except OSError as exc:
        if 'WinError 2' in str(exc):
            raise RuntimeError(
                "Audio engine binary not found (WinError 2). Ensure ffmpeg/ffprobe are accessible."
            ) from exc
        raise


_configure_audio_binaries()

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
    audio = _load_audio_segment(input_path)
    source_duration_ms = len(audio)
    
    # 1. Trimming
    start_time = tool_params.get('start')
    end_time = tool_params.get('end')
    try:
        start_ms = max(0, int(float(start_time or 0) * 1000))
        end_ms = int(float(end_time) * 1000) if end_time not in (None, "", "None") else source_duration_ms
    except (ValueError, TypeError):
        raise ValueError("Invalid trim range values.")

    end_ms = min(source_duration_ms, end_ms)
    if end_ms <= start_ms:
        raise ValueError("End time must be greater than start time.")
    audio = audio[start_ms:end_ms]

    # 2. Volume
    try:
        volume_level = float(tool_params.get('volume', 100)) # percentage
        volume_level = min(max(volume_level, 0), 300)
        if volume_level != 100:
            if volume_level > 0:
                gain_db = 20 * math.log10(volume_level / 100.0)
                audio = audio.apply_gain(gain_db)
            else:
                audio = audio - 120 # Mute
    except (ValueError, TypeError):
        raise ValueError("Invalid volume value.")

    # 3. Speed (preserve original voice pitch as much as possible)
    try:
        speed = float(tool_params.get('speed', 1.0))
        speed = min(max(speed, 0.5), 2.5)
        if speed != 1.0:
            # pydub.effects.speedup keeps voice closer to original than raw frame-rate shifting.
            if speed > 1.0:
                audio = speedup(audio, playback_speed=speed, chunk_size=120, crossfade=20)
            else:
                # Slowing down with simple resampling is fallback behavior.
                # It is only used when user explicitly asks for <1.0 speed.
                new_sample_rate = int(audio.frame_rate * speed)
                audio = audio._spawn(audio.raw_data, overrides={'frame_rate': new_sample_rate})
                audio = audio.set_frame_rate(audio.frame_rate)
    except (ValueError, TypeError):
        raise ValueError("Invalid speed value.")

    # 4. Pitch
    try:
        pitch = float(tool_params.get('pitch', 0))
        pitch = min(max(pitch, -12), 12)
        if pitch != 0:
            # Shift pitch by changing sample rate
            new_sample_rate = int(audio.frame_rate * (2.0 ** (pitch / 12.0)))
            audio = audio._spawn(audio.raw_data, overrides={'frame_rate': new_sample_rate})
            audio = audio.set_frame_rate(44100)
    except (ValueError, TypeError):
        raise ValueError("Invalid pitch value.")

    # 5. Equalizer (Presets)
    preset = tool_params.get('preset', 'none')
    if preset != 'none':
        if preset == 'full-bass':
            preset = 'bass-boost'
        if preset == 'full-treble':
            preset = 'treble-boost'
        if preset == 'bass-boost':
            audio = audio.low_pass_filter(250).apply_gain(6)
        elif preset == 'treble-boost':
            audio = audio.high_pass_filter(5000).apply_gain(6)
        elif preset == 'classic':
            audio = audio.high_pass_filter(1000).low_pass_filter(4000).apply_gain(3)
        elif preset == 'dance':
            audio = audio.low_pass_filter(200).apply_gain(4).high_pass_filter(6000).apply_gain(4)
        elif preset == 'club':
            audio = audio.low_pass_filter(120).apply_gain(5).high_pass_filter(5000).apply_gain(2)
        elif preset == 'pop':
            audio = audio.high_pass_filter(1000).apply_gain(2)
        elif preset == 'rock':
            audio = audio.low_pass_filter(150).apply_gain(3).high_pass_filter(3000).apply_gain(3)

    # 6. Fade in/out
    try:
        fade_in_ms = int(float(tool_params.get('fade_in', 0)) * 1000)
        fade_out_ms = int(float(tool_params.get('fade_out', 0)) * 1000)
        max_fade = max(0, (len(audio) // 2) - 1)
        fade_in_ms = min(max(fade_in_ms, 0), max_fade)
        fade_out_ms = min(max(fade_out_ms, 0), max_fade)
        if fade_in_ms > 0:
            audio = audio.fade_in(fade_in_ms)
        if fade_out_ms > 0:
            audio = audio.fade_out(fade_out_ms)
    except (ValueError, TypeError):
        raise ValueError("Invalid fade values.")
    # 5. Reverse
    if tool_params.get('reverse') == 'true':
        audio = audio.reverse()

    target_format = tool_params.get('format', 'mp3')
    if target_format not in {'mp3', 'wav', 'ogg', 'm4a', 'flac'}:
        raise ValueError("Unsupported output format.")

    bitrate = tool_params.get('bitrate') or '320k'
    output_path = get_output_path(original_name, target_format, '_processed')
    export_kwargs = {}
    if target_format == 'mp3':
        export_kwargs['bitrate'] = bitrate
    audio.export(output_path, format=target_format if target_format != 'm4a' else 'mp4', **export_kwargs)
    return output_path

def merge_audios(input_paths, original_name, target_format='mp3'):
    combined = AudioSegment.empty()
    for path in input_paths:
        track = _load_audio_segment(path)
        combined += track
    
    output_path = get_output_path(original_name, target_format, '_merged')
    combined.export(output_path, format=target_format)
    return output_path


def extract_audio_from_video(input_path, original_name, target_format='mp3',
                             extract_mode='full', start=0, end=''):
    video_audio = _load_audio_segment(input_path)
    if extract_mode == 'range':
        try:
            start_ms = max(0, int(float(start or 0) * 1000))
            end_ms = int(float(end) * 1000) if end not in (None, '', 'None') else len(video_audio)
        except (TypeError, ValueError):
            raise ValueError("Invalid extract range values.")
        end_ms = min(end_ms, len(video_audio))
        if end_ms <= start_ms:
            raise ValueError("End time must be greater than start time for range extraction.")
        video_audio = video_audio[start_ms:end_ms]

    output_path = get_output_path(original_name, target_format, '_extracted')
    video_audio.export(output_path, format=target_format)
    return output_path
