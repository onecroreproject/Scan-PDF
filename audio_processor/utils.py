import os
import json
import logging
import subprocess
import tempfile
import uuid
import re
from pathlib import Path
from converter.media_binaries import resolve_ffmpeg_paths, ensure_ffmpeg_configured

logger = logging.getLogger(__name__)

# Ensure FFmpeg is configured
ensure_ffmpeg_configured()

SUPPORTED_AUDIO_EXTENSIONS = {'.mp3', '.wav', '.ogg', '.m4a', '.flac', '.aac', '.wma', '.webm'}
SUPPORTED_VIDEO_EXTENSIONS = {'.mp4', '.mov', '.avi', '.mkv', '.webm', '.m4v', '.wmv', '.flv', '.3gp'}
SUPPORTED_OUTPUT_FORMATS = {'mp3', 'wav', 'ogg', 'm4a', 'flac'}

def get_ffmpeg_binaries():
    """Returns (ffmpeg_path, ffprobe_path) using the project's resolver."""
    ffmpeg_bin, ffprobe_bin = resolve_ffmpeg_paths()
    return ffmpeg_bin or 'ffmpeg', ffprobe_bin or 'ffprobe'


def probe_media(file_path):
    """
    Probe a media file using ffprobe.
    Returns a dict with:
      - duration: float
      - has_audio: bool
      - audio_codec: str or None
      - sample_rate: int or None
      - channels: int or None
      - has_video: bool
      - video_codec: str or None
      - width: int or None
      - height: int or None
    """
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        raise ValueError("Uploaded file is empty or missing.")

    _, ffprobe_bin = get_ffmpeg_binaries()
    # If ffprobe binary is unavailable, fall back to using ffmpeg -i parsing.
    if not ffprobe_bin or not os.path.exists(ffprobe_bin):
        ffmpeg_bin, _ = get_ffmpeg_binaries()
        fallback_cmd = [ffmpeg_bin, '-i', str(file_path)]
        try:
            proc = subprocess.run(
                fallback_cmd,
                capture_output=True,
                text=True,
                check=False,
                timeout=30
            )
            stderr = proc.stderr
            has_audio = bool(re.search(r'Audio:\s', stderr, re.IGNORECASE))
            # Extract duration if possible.
            dur_match = re.search(r'Duration: (\d{2}):(\d{2}):(\d{2}\.\d{2})', stderr)
            duration = 0.0
            if dur_match:
                h, m, s = dur_match.groups()
                duration = int(h) * 3600 + int(m) * 60 + float(s)
            return {
                'duration': duration,
                'has_audio': has_audio,
                'audio_codec': None,
                'sample_rate': None,
                'channels': None,
                'has_video': True,
                'video_codec': None,
                'width': None,
                'height': None,
                'size': os.path.getsize(file_path),
            }
        except subprocess.TimeoutExpired:
            raise RuntimeError("Media probing timed out.")
        except Exception as e:
            logger.error(f"ffmpeg fallback probe failed for {file_path}: {e}")
            raise ValueError("Invalid or corrupted media file.")
    
    cmd = [
        ffprobe_bin,
        '-v', 'quiet',
        '-print_format', 'json',
        '-show_format',
        '-show_streams',
        str(file_path)
    ]

    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            timeout=30
        )
        data = json.loads(proc.stdout)
    except subprocess.TimeoutExpired:
        raise RuntimeError("Media probing timed out.")
    except (subprocess.CalledProcessError, json.JSONDecodeError, Exception) as e:
        logger.error(f"ffprobe failed for {file_path}: {e}")
        raise ValueError("Invalid or corrupted media file.")

    streams = data.get('streams', [])
    fmt = data.get('format', {})

    has_audio = False
    audio_codec = None
    sample_rate = None
    channels = None
    has_video = False
    video_codec = None
    width = None
    height = None

    for s in streams:
        codec_type = s.get('codec_type')
        if codec_type == 'audio' and not has_audio:
            has_audio = True
            audio_codec = s.get('codec_name')
            try:
                sample_rate = int(s.get('sample_rate', 0))
            except (ValueError, TypeError):
                sample_rate = None
            try:
                channels = int(s.get('channels', 0))
            except (ValueError, TypeError):
                channels = None
        elif codec_type == 'video' and not has_video:
            # Check it's not an attached pic
            if s.get('disposition', {}).get('attached_pic', 0) == 0:
                has_video = True
                video_codec = s.get('codec_name')
                width = s.get('width')
                height = s.get('height')

    duration = 0.0
    # Try format duration first, then stream duration
    try:
        duration = float(fmt.get('duration', 0.0))
    except (ValueError, TypeError):
        duration = 0.0

    if duration <= 0:
        for s in streams:
            try:
                dur = float(s.get('duration', 0.0))
                if dur > duration:
                    duration = dur
            except (ValueError, TypeError):
                pass

    return {
        'duration': max(0.0, duration),
        'has_audio': has_audio,
        'audio_codec': audio_codec,
        'sample_rate': sample_rate,
        'channels': channels,
        'has_video': has_video,
        'video_codec': video_codec,
        'width': width,
        'height': height,
        'size': os.path.getsize(file_path),
    }


def get_safe_output_path(original_name, target_format, suffix='_processed'):
    from converter.utils import ensure_media_dirs
    _, output_dir = ensure_media_dirs()
    base_name = os.path.splitext(os.path.basename(original_name))[0]
    # Clean base_name
    safe_base = "".join(c for c in base_name if c.isalnum() or c in ('-', '_')).strip() or 'audio'
    unique_id = uuid.uuid4().hex[:8]
    filename = f"{safe_base}{suffix}_{unique_id}.{target_format.lower()}"
    return os.path.join(output_dir, filename)


def extract_audio_from_video(input_path, original_name, target_format='mp3', extract_mode='full', start=0, end=''):
    """
    Direct FFmpeg extraction from video to audio.
    Prevents MoviePy buffer errors on videos with no audio or corrupted frames.
    """
    target_format = (target_format or 'mp3').lower().strip()
    if target_format not in SUPPORTED_OUTPUT_FORMATS:
        target_format = 'mp3'

    # Probe media file and log detailed information for diagnostics
    try:
        probe = probe_media(input_path)
        logger.debug(f"Probe result for {input_path}: {probe}")
    except Exception as probe_err:
        logger.error(f"Failed to probe video file {input_path}: {probe_err}")
        raise ValueError(f"Video probe failed: {probe_err}")

    if not probe.get('has_audio'):
        # No audio stream detected – inform caller
        raise ValueError("No audio track found in the uploaded video.")

    duration = probe['duration']
    ffmpeg_bin, _ = get_ffmpeg_binaries()

    start_sec = 0.0
    end_sec = duration

    if extract_mode == 'range':
        try:
            start_sec = max(0.0, float(start or 0.0))
        except (ValueError, TypeError):
            start_sec = 0.0

        try:
            if end not in (None, '', 'None'):
                end_sec = float(end)
            else:
                end_sec = duration
        except (ValueError, TypeError):
            end_sec = duration

        if duration > 0:
            end_sec = min(end_sec, duration)

        if end_sec <= start_sec:
            raise ValueError("End time must be greater than start time for range extraction.")

    output_path = get_safe_output_path(original_name, target_format, suffix='_extracted')

    cmd = [ffmpeg_bin, '-y']

    # Fast seek before input
    if start_sec > 0:
        cmd.extend(['-ss', str(start_sec)])

    cmd.extend(['-i', str(input_path)])

    # Duration to extract
    if extract_mode == 'range' and end_sec > start_sec:
        range_duration = end_sec - start_sec
        cmd.extend(['-t', str(range_duration)])

    # Map audio only
    cmd.extend(['-vn'])

    # Codec and quality settings
    if target_format == 'mp3':
        cmd.extend(['-c:a', 'libmp3lame', '-b:a', '192k'])
    elif target_format == 'wav':
        cmd.extend(['-c:a', 'pcm_s16le'])
    elif target_format == 'ogg':
        cmd.extend(['-c:a', 'libvorbis', '-q:a', '5'])
    elif target_format in ('m4a', 'aac'):
        cmd.extend(['-c:a', 'aac', '-b:a', '192k'])
    elif target_format == 'flac':
        cmd.extend(['-c:a', 'flac'])
    else:
        cmd.extend(['-c:a', 'libmp3lame', '-b:a', '192k'])

    cmd.append(str(output_path))

    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            timeout=180
        )
    except subprocess.TimeoutExpired:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        raise RuntimeError("Audio extraction timed out.")
    except subprocess.CalledProcessError as e:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        logger.error(f"FFmpeg extract failed: {e.stderr}")
        raise RuntimeError(f"Audio extraction failed: {e.stderr.strip() if e.stderr else str(e)}")

    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError("Extracted audio output is empty or missing.")

    # Validate output probe
    out_probe = probe_media(output_path)
    if not out_probe['has_audio'] or out_probe['duration'] <= 0:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        raise RuntimeError("Extracted audio is corrupted or unplayable.")

    return output_path


def merge_audios(input_paths, original_name, target_format='mp3'):
    """
    Merge multiple audio files with diverse codecs/sample rates using FFmpeg concat filter.
    """
    if not input_paths or len(input_paths) < 2:
        raise ValueError("At least 2 audio files are required to merge.")

    target_format = (target_format or 'mp3').lower().strip()
    if target_format not in SUPPORTED_OUTPUT_FORMATS:
        target_format = 'mp3'

    ffmpeg_bin, _ = get_ffmpeg_binaries()
    cmd = [ffmpeg_bin, '-y']

    for p in input_paths:
        if not os.path.exists(p) or os.path.getsize(p) == 0:
            raise ValueError("One or more audio tracks are invalid or empty.")
        # Probe to ensure audio stream exists
        probe = probe_media(p)
        if not probe['has_audio']:
            raise ValueError(f"File '{os.path.basename(p)}' does not contain a valid audio stream.")
        cmd.extend(['-i', str(p)])

    # Construct concat filter: [0:a][1:a]...concat=n=N:v=0:a=1[outa]
    n = len(input_paths)
    filter_inputs = "".join(f"[{i}:a]" for i in range(n))
    filter_str = f"{filter_inputs}concat=n={n}:v=0:a=1[outa]"

    cmd.extend(['-filter_complex', filter_str, '-map', '[outa]'])

    if target_format == 'mp3':
        cmd.extend(['-c:a', 'libmp3lame', '-b:a', '192k'])
    elif target_format == 'wav':
        cmd.extend(['-c:a', 'pcm_s16le'])
    elif target_format == 'ogg':
        cmd.extend(['-c:a', 'libvorbis', '-q:a', '5'])
    elif target_format in ('m4a', 'aac'):
        cmd.extend(['-c:a', 'aac', '-b:a', '192k'])
    elif target_format == 'flac':
        cmd.extend(['-c:a', 'flac'])

    output_path = get_safe_output_path(original_name, target_format, suffix='_merged')
    cmd.append(str(output_path))

    try:
        subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            timeout=240
        )
    except subprocess.TimeoutExpired:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        raise RuntimeError("Audio merge timed out.")
    except subprocess.CalledProcessError as e:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        logger.error(f"FFmpeg merge failed: {e.stderr}")
        raise RuntimeError(f"Audio merge failed: {e.stderr.strip() if e.stderr else str(e)}")

    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError("Merged audio output is empty.")

    out_probe = probe_media(output_path)
    if not out_probe['has_audio'] or out_probe['duration'] <= 0:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        raise RuntimeError("Merged audio is invalid or corrupted.")

    return output_path


def process_audio(input_path, original_name, tool_params):
    """
    Process audio: trim, volume adjustment, fade in/out, format conversion.
    Direct FFmpeg pipeline for rock-solid stability and zero buffer crashes.
    """
    if not os.path.exists(input_path) or os.path.getsize(input_path) == 0:
        raise ValueError("Audio file is missing or empty.")

    probe = probe_media(input_path)
    if not probe['has_audio']:
        raise ValueError("Uploaded file does not contain a valid audio stream.")

    duration = probe['duration']
    ffmpeg_bin, _ = get_ffmpeg_binaries()

    # 1. Trimming parameters
    start_time = tool_params.get('start')
    end_time = tool_params.get('end')
    try:
        start_sec = max(0.0, float(start_time or 0.0))
    except (ValueError, TypeError):
        start_sec = 0.0

    try:
        if end_time not in (None, '', 'None', '0', 0):
            end_sec = float(end_time)
        else:
            end_sec = duration
    except (ValueError, TypeError):
        end_sec = duration

    if duration > 0:
        end_sec = min(end_sec, duration)

    if end_sec <= start_sec:
        raise ValueError("End time must be greater than start time.")

    # 2. Volume parameter (percentage: 0 to 300)
    try:
        vol_pct = float(tool_params.get('volume', 100))
        vol_pct = max(0.0, min(vol_pct, 300.0))
        vol_factor = vol_pct / 100.0
    except (ValueError, TypeError):
        vol_factor = 1.0

    # 3. Fade in / Fade out
    try:
        fade_in_sec = max(0.0, float(tool_params.get('fade_in', 0) or 0))
    except (ValueError, TypeError):
        fade_in_sec = 0.0

    try:
        fade_out_sec = max(0.0, float(tool_params.get('fade_out', 0) or 0))
    except (ValueError, TypeError):
        fade_out_sec = 0.0

    clip_len = end_sec - start_sec
    fade_in_sec = min(fade_in_sec, clip_len / 2.0)
    fade_out_sec = min(fade_out_sec, clip_len / 2.0)

    # 4. Target format
    target_format = (tool_params.get('format') or 'mp3').lower().strip()
    if target_format not in SUPPORTED_OUTPUT_FORMATS:
        target_format = 'mp3'

    # Build audio filters
    filters = []
    if vol_factor != 1.0:
        filters.append(f"volume={vol_factor:.2f}")

    if fade_in_sec > 0:
        filters.append(f"afade=t=in:ss=0:d={fade_in_sec:.2f}")

    if fade_out_sec > 0:
        fade_out_start = max(0.0, clip_len - fade_out_sec)
        filters.append(f"afade=t=out:st={fade_out_start:.2f}:d={fade_out_sec:.2f}")

    # Build command
    cmd = [ffmpeg_bin, '-y']

    if start_sec > 0:
        cmd.extend(['-ss', str(start_sec)])

    cmd.extend(['-i', str(input_path)])

    if clip_len > 0 and (start_sec > 0 or end_sec < duration):
        cmd.extend(['-t', str(clip_len)])

    cmd.extend(['-vn'])

    if filters:
        cmd.extend(['-af', ",".join(filters)])

    # Output encoding
    if target_format == 'mp3':
        bitrate = tool_params.get('bitrate') or '192k'
        cmd.extend(['-c:a', 'libmp3lame', '-b:a', str(bitrate)])
    elif target_format == 'wav':
        cmd.extend(['-c:a', 'pcm_s16le'])
    elif target_format == 'ogg':
        cmd.extend(['-c:a', 'libvorbis', '-q:a', '5'])
    elif target_format in ('m4a', 'aac'):
        cmd.extend(['-c:a', 'aac', '-b:a', '192k'])
    elif target_format == 'flac':
        cmd.extend(['-c:a', 'flac'])

    output_path = get_safe_output_path(original_name, target_format, suffix='_edited')
    cmd.append(str(output_path))

    try:
        subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            timeout=180
        )
    except subprocess.TimeoutExpired:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        raise RuntimeError("Audio editing operation timed out.")
    except subprocess.CalledProcessError as e:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        logger.error(f"FFmpeg edit failed: {e.stderr}")
        raise RuntimeError(f"Audio processing failed: {e.stderr.strip() if e.stderr else str(e)}")

    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError("Edited audio output is empty.")

    out_probe = probe_media(output_path)
    if not out_probe['has_audio'] or out_probe['duration'] <= 0:
        if os.path.exists(output_path):
            try: os.remove(output_path)
            except OSError: pass
        raise RuntimeError("Processed audio is invalid or corrupted.")

    return output_path
