"""
FFmpeg command builders and execution helpers.
Optimized for Hostinger VPS with limited RAM.
"""
import os
import re
import uuid
import shutil
import subprocess
import logging
from pathlib import Path
from django.conf import settings

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════
# PATHS
# ═══════════════════════════════════════════════════════════════

def get_ffmpeg_path():
    """Return FFmpeg binary path."""
    if hasattr(settings, 'FFMPEG_PATH') and os.path.exists(settings.FFMPEG_PATH):
        return settings.FFMPEG_PATH
    # Fallback to system PATH
    ffmpeg = shutil.which('ffmpeg')
    if ffmpeg:
        return ffmpeg
    raise RuntimeError("FFmpeg not found. Install FFmpeg or set FFMPEG_PATH in settings.")


def get_ffprobe_path():
    """Return FFprobe binary path."""
    if hasattr(settings, 'FFPROBE_PATH') and os.path.exists(settings.FFPROBE_PATH):
        return settings.FFPROBE_PATH
    ffprobe = shutil.which('ffprobe')
    if ffprobe:
        return ffprobe
    raise RuntimeError("FFprobe not found. Install FFmpeg or set FFPROBE_PATH in settings.")

# ═══════════════════════════════════════════════════════════════
# VIDEO INFO
# ═══════════════════════════════════════════════════════════════

def get_video_info(video_path):
    """Extract video metadata using ffprobe."""
    ffprobe = get_ffprobe_path()
    cmd = [
        ffprobe,
        '-v', 'error',
        '-select_streams', 'v:0',
        '-show_entries', 'stream=width,height,duration,r_frame_rate,codec_name,pix_fmt',
        '-show_entries', 'format=duration,size,bit_rate',
        '-of', 'default=noprint_wrappers=1',
        video_path
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        if result.returncode != 0:
            logger.error(f"ffprobe error: {result.stderr}")
            return {}

        info = {}
        for line in result.stdout.strip().split('\n'):
            if '=' in line:
                key, val = line.split('=', 1)
                info[key.strip()] = val.strip()

        # Parse resolution
        info['width'] = int(info.get('width', 0))
        info['height'] = int(info.get('height', 0))

        # Parse duration
        try:
            info['duration'] = float(info.get('duration', 0))
        except (ValueError, TypeError):
            info['duration'] = 0

        # Parse frame rate
        fps_str = info.get('r_frame_rate', '0/1')
        if '/' in fps_str:
            num, den = fps_str.split('/')
            try:
                info['fps'] = float(num) / float(den)
            except (ValueError, ZeroDivisionError):
                info['fps'] = 0
        else:
            try:
                info['fps'] = float(fps_str)
            except ValueError:
                info['fps'] = 0

        # Parse bitrate
        try:
            info['bitrate'] = int(info.get('bit_rate', 0))
        except (ValueError, TypeError):
            info['bitrate'] = 0

        return info
    except subprocess.TimeoutExpired:
        logger.error("ffprobe timeout")
        return {}
    except Exception as e:
        logger.error(f"ffprobe exception: {e}")
        return {}

# ═══════════════════════════════════════════════════════════════
# RESOLUTION MAP
# ═══════════════════════════════════════════════════════════════

RESOLUTION_MAP = {
    '360p':  {'width': 640,  'height': 360},
    '480p':  {'width': 854,  'height': 480},
    '720p':  {'width': 1280, 'height': 720},
    '1080p': {'width': 1920, 'height': 1080},
    '2k':    {'width': 2560, 'height': 1440},
    '4k':    {'width': 3840, 'height': 2160},
}

QUALITY_PRESETS = {
    'low':    {'crf': 28, 'preset': 'ultrafast', 'bufsize': '1M', 'maxrate': '1M'},
    'medium': {'crf': 23, 'preset': 'fast',      'bufsize': '2M', 'maxrate': '2M'},
    'high':   {'crf': 18, 'preset': 'medium',    'bufsize': '5M', 'maxrate': '5M'},
}

# ═══════════════════════════════════════════════════════════════
# CORE RUNNER
# ═══════════════════════════════════════════════════════════════

def run_ffmpeg(args, timeout=600):
    """
    Execute FFmpeg safely with subprocess.
    args: list of arguments (without 'ffmpeg' binary)
    timeout: max seconds
    """
    ffmpeg = get_ffmpeg_path()
    cmd = [ffmpeg] + args
    logger.info(f"FFmpeg command: {' '.join(cmd)}")

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        stdout, stderr = proc.communicate(timeout=timeout)
        if proc.returncode != 0:
            logger.error(f"FFmpeg error (code {proc.returncode}): {stderr}")
            raise RuntimeError(f"FFmpeg failed: {stderr[:500]}")
        return stdout, stderr
    except subprocess.TimeoutExpired:
        proc.kill()
        raise RuntimeError("FFmpeg processing timed out")
    except Exception as e:
        raise RuntimeError(f"FFmpeg execution error: {e}")


# ═══════════════════════════════════════════════════════════════
# MODULE 1: VIDEO CONVERTER
# ═══════════════════════════════════════════════════════════════

def build_convert_command(input_path, output_path, output_format, options=None):
    """Build FFmpeg command for video conversion."""
    options = options or {}
    resolution = options.get('resolution', 'original')
    quality = options.get('quality', 'medium')
    fps = options.get('fps')
    codec = options.get('codec')
    bitrate = options.get('bitrate')

    preset = QUALITY_PRESETS.get(quality, QUALITY_PRESETS['medium'])

    # Select codec
    fmt = output_format.lower()
    if not codec:
        if fmt == 'webm':
            codec = 'libvpx-vp9'
        elif fmt == 'avi':
            codec = 'mpeg4'
        else:
            codec = 'libx264'

    audio_codec = 'aac'
    if fmt == 'webm':
        audio_codec = 'libopus'
    elif fmt == 'avi':
        audio_codec = 'libmp3lame'

    cmd = [
        '-y',  # overwrite
        '-hide_banner',
        '-loglevel', 'error',
        '-i', input_path,
    ]

    # Threads for multi-core
    cmd += ['-threads', '0']

    # Resolution scaling
    if resolution != 'original' and resolution in RESOLUTION_MAP:
        res = RESOLUTION_MAP[resolution]
        # Use scale with force_original_aspect_ratio to avoid distortion
        cmd += ['-vf', f"scale={res['width']}:{res['height']}:force_original_aspect_ratio=decrease,pad={res['width']}:{res['height']}:(ow-iw)/2:(oh-ih)/2:black"]

    # FPS
    if fps:
        cmd += ['-r', str(fps)]

    # Video codec
    cmd += ['-c:v', codec]
    if codec in ('libx264', 'libx265'):
        cmd += ['-preset', preset['preset'], '-crf', str(preset['crf'])]
    elif codec == 'libvpx-vp9':
        cmd += ['-crf', str(preset['crf']), '-b:v', '0']  # VP9 CRF mode

    # Bitrate control
    if bitrate:
        cmd += ['-b:v', bitrate]
    else:
        cmd += ['-maxrate', preset['maxrate'], '-bufsize', preset['bufsize']]

    # Audio
    cmd += ['-c:a', audio_codec, '-b:a', '128k']

    # Fast start for web
    if fmt == 'mp4':
        cmd += ['-movflags', '+faststart']

    # Pixel format
    cmd += ['-pix_fmt', 'yuv420p']

    cmd.append(output_path)
    return cmd


# ═══════════════════════════════════════════════════════════════
# MODULE 2: IMAGE TO VIDEO
# ═══════════════════════════════════════════════════════════════

def build_image_to_video_command(image_paths, output_path, options=None):
    """
    Build and execute FFmpeg command for creating video from images.
    Supports per-image durations, audio trimming with fade/volume, and transitions.
    Handles audio/video duration mismatch (trim, loop, or silence).
    Returns (output_path, temp_files_list).
    """
    options = options or {}
    durations = options.get('durations', [3.0] * len(image_paths))
    fps = options.get('fps', 30)
    resolution = options.get('resolution', '1080p')
    aspect_ratio = options.get('aspect_ratio', '16:9')  # 16:9, 9:16, 1:1
    transition = options.get('transition', 'none')
    audio_path = options.get('audio_path')
    audio_trim_start = options.get('audio_trim_start', 0.0)
    audio_trim_end = options.get('audio_trim_end')
    audio_volume = options.get('audio_volume', 1.0)  # 0.0 to 2.0
    audio_fade_in = options.get('audio_fade_in', 0.0)  # seconds
    audio_fade_out = options.get('audio_fade_out', 0.0)  # seconds
    audio_behavior = options.get('audio_behavior', 'trim')  # trim, loop, silence
    filename = options.get('filename', 'slideshow.mp4')

    # Calculate target dimensions based on resolution and aspect ratio
    res = RESOLUTION_MAP.get(resolution, RESOLUTION_MAP['1080p'])
    base_width, base_height = res['width'], res['height']
    
    # Apply aspect ratio
    if aspect_ratio == '1:1':
        width = height = min(base_width, base_height)
    elif aspect_ratio == '9:16':
        # Vertical video - swap dimensions and scale
        height = base_width
        width = int(base_width * 9 / 16)
    else:  # 16:9 default
        width, height = base_width, base_height

    temp_files = []
    trans_dur = 1.0

    # Calculate total video duration
    video_duration = sum(float(durations[i]) if i < len(durations) else 3.0 
                        for i in range(len(image_paths)))
    # Account for transitions (they overlap)
    if transition in ('slideleft', 'slideright', 'slideup', 'slidedown', 
                       'wipeleft', 'wiperight', 'fade') and len(image_paths) > 1:
        video_duration -= trans_dur * (len(image_paths) - 1)
    
    # Calculate required loop count if audio is longer than video
    loop_count = 1
    if audio_path:
        audio_info = get_video_info(audio_path)
        if audio_info.get('duration'):
            audio_total_duration = audio_info['duration']
            if audio_trim_end and audio_trim_start:
                audio_trim_duration = audio_trim_end - audio_trim_start
            elif audio_trim_start > 0:
                audio_trim_duration = audio_total_duration - audio_trim_start
            else:
                audio_trim_duration = audio_total_duration
            
            if audio_trim_duration > video_duration:
                loop_count = int(math.ceil(audio_trim_duration / video_duration))

    # ═══════════════════════════════════════════════════════════════
    # STEP 1: Generate individual video segments from each image
    # ═══════════════════════════════════════════════════════════════
    segments = []
    for i, img_path in enumerate(image_paths):
        dur = float(durations[i]) if i < len(durations) else 3.0
        seg_path = os.path.join(os.path.dirname(output_path), f"seg_{uuid.uuid4().hex}.mp4")
        vf = f"scale={width}:{height}:force_original_aspect_ratio=decrease,pad={width}:{height}:(ow-iw)/2:(oh-ih)/2:black,format=yuv420p,fps={fps}"

        if transition == 'zoom':
            total_frames = int(dur * fps)
            vf += f",zoompan=z='min(zoom+0.0015,1.5)':d={total_frames}:s={width}x{height}:x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)'"

        cmd = [
            '-y', '-hide_banner', '-loglevel', 'error',
            '-loop', '1', '-i', img_path,
            '-vf', vf,
            '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
            '-t', str(dur), '-an', '-pix_fmt', 'yuv420p',
            seg_path
        ]
        run_ffmpeg(cmd, timeout=300)
        segments.append(seg_path)
        temp_files.append(seg_path)

    video_output = None

    # ═══════════════════════════════════════════════════════════════
    # STEP 2: Merge segments with or without transitions
    # ═══════════════════════════════════════════════════════════════
    if len(segments) == 1:
        video_output = segments[0]
    elif transition in ('slideleft', 'slideright', 'slideup', 'slidedown', 'wipeleft', 'wiperight', 'fade'):
        # Use xfade complex filter for transitions
        trans_map = {
            'slideleft': 'slideleft', 'slideright': 'slideright',
            'slideup': 'slideup', 'slidedown': 'slidedown',
            'wipeleft': 'wipeleft', 'wiperight': 'wiperight',
            'fade': 'fade',
        }
        tname = trans_map.get(transition, 'fade')

        # Ensure each duration is at least trans_dur to prevent negative offsets
        safe_durations = [max(float(d), trans_dur) for d in durations]

        inputs = []
        for seg in segments:
            inputs += ['-i', seg]

        filter_parts = []
        last_label = '0:v'
        for i in range(1, len(segments)):
            offset = sum(safe_durations[j] for j in range(i)) - trans_dur
            out_label = f"tmp{i}" if i < len(segments) - 1 else "outv"
            filter_parts.append(
                f"[{last_label}][{i}:v]xfade=transition={tname}:duration={trans_dur}:offset={offset}[{out_label}]"
            )
            last_label = out_label

        cmd = ['-y', '-hide_banner', '-loglevel', 'error']
        cmd += inputs
        cmd += ['-filter_complex', ';'.join(filter_parts)]
        cmd += ['-map', '[outv]', '-c:v', 'libx264', '-preset', 'fast', '-crf', '23', '-an', '-pix_fmt', 'yuv420p']
        cmd.append(output_path)
        run_ffmpeg(cmd, timeout=1800)
        video_output = output_path
    else:
        # Simple concat demuxer (no transitions, or fade handled per-segment)
        concat_file = os.path.join(os.path.dirname(output_path), f"concat_{uuid.uuid4().hex}.txt")
        with open(concat_file, 'w') as f:
            for seg in segments:
                f.write(f"file '{os.path.abspath(seg)}'\n")
        temp_files.append(concat_file)

        cmd = [
            '-y', '-hide_banner', '-loglevel', 'error',
            '-f', 'concat', '-safe', '0', '-i', concat_file,
            '-c', 'copy', '-movflags', '+faststart',
            output_path
        ]
        run_ffmpeg(cmd, timeout=600)
        video_output = output_path

    # ═══════════════════════════════════════════════════════════════
    # STEP 3: Process and add audio if provided
    # ═══════════════════════════════════════════════════════════════
    if audio_path and os.path.exists(audio_path):
        # Calculate trimmed audio duration
        trimmed_duration = None
        if audio_trim_end and audio_trim_end > audio_trim_start:
            trimmed_duration = audio_trim_end - audio_trim_start
        elif audio_trim_start > 0:
            # Get audio info to calculate duration
            audio_info = get_video_info(audio_path)
            if audio_info.get('duration'):
                trimmed_duration = audio_info['duration'] - audio_trim_start

        # Build audio filter chain
        audio_filters = []
        if audio_volume != 1.0:
            audio_filters.append(f"volume={audio_volume}")
        
        # Add fade in/out
        if trimmed_duration:
            if audio_fade_in > 0:
                audio_filters.append(f"afade=t=in:ss=0:d={min(audio_fade_in, trimmed_duration/2)}")
            if audio_fade_out > 0:
                fade_start = max(0, trimmed_duration - audio_fade_out)
                audio_filters.append(f"afade=t=out:st={fade_start}:d={min(audio_fade_out, trimmed_duration/2)}")

        # Process audio with trimming and filters
        processed_audio = os.path.join(os.path.dirname(output_path), f"audio_proc_{uuid.uuid4().hex}.m4a")
        
        # Use input seeking for accurate trimming (-ss before -i)
        # When using input seeking, -t (duration) is more accurate than -to
        audio_cmd = ['-y', '-hide_banner', '-loglevel', 'error']
        if audio_trim_start > 0:
            audio_cmd += ['-ss', str(audio_trim_start)]
        if audio_trim_end:
            # Calculate duration from start to end
            trim_duration = audio_trim_end - audio_trim_start
            audio_cmd += ['-t', str(trim_duration)]
        audio_cmd += ['-i', audio_path]
        
        if audio_filters:
            audio_cmd += ['-af', ','.join(audio_filters)]
        
        # Handle duration mismatch
        if trimmed_duration and audio_behavior == 'loop' and trimmed_duration < video_duration:
            # Loop audio to match video duration
            loop_count = int(video_duration / trimmed_duration) + 1
            audio_cmd += ['-stream_loop', str(loop_count)]
        
        audio_cmd += ['-c:a', 'aac', '-b:a', '192k', '-vn', processed_audio]
        run_ffmpeg(audio_cmd, timeout=300)
        temp_files.append(processed_audio)

        # If audio is shorter than video and behavior is 'silence', pad with silence
        if trimmed_duration and trimmed_duration < video_duration and audio_behavior == 'silence':
            padded_audio = os.path.join(os.path.dirname(output_path), f"audio_padded_{uuid.uuid4().hex}.m4a")
            pad_cmd = [
                '-y', '-hide_banner', '-loglevel', 'error',
                '-i', processed_audio,
                '-af', f'apad=pad_dur={video_duration - trimmed_duration}',
                '-c:a', 'aac', '-b:a', '192k',
                '-t', str(video_duration),
                padded_audio
            ]
            run_ffmpeg(pad_cmd, timeout=60)
            temp_files.append(processed_audio)  # Mark original for deletion
            processed_audio = padded_audio

        # Merge video + audio
        final_output = os.path.join(os.path.dirname(output_path), f"final_{uuid.uuid4().hex}.mp4")
        
        merge_cmd = [
            '-y', '-hide_banner', '-loglevel', 'error',
            '-i', video_output,
            '-i', processed_audio,
            '-c:v', 'copy',
            '-c:a', 'aac', '-b:a', '192k',
            '-movflags', '+faststart',
        ]
        
        # Handle which stream determines duration
        if audio_behavior == 'trim' and trimmed_duration and trimmed_duration < video_duration:
            # Trim video to match audio
            merge_cmd += ['-t', str(trimmed_duration)]
        elif loop_count > 1:
            # Loop video to match audio duration
            merge_cmd += ['-stream_loop', str(loop_count), '-shortest']
        else:
            # Shortest ensures video ends when audio ends (or vice versa)
            merge_cmd += ['-shortest']
        
        merge_cmd.append(final_output)
        run_ffmpeg(merge_cmd, timeout=300)

        if video_output != output_path and video_output not in temp_files:
            temp_files.append(video_output)
        temp_files.append(processed_audio)
        video_output = final_output

    # Move final output to expected path if needed
    if video_output != output_path:
        if os.path.exists(output_path):
            os.remove(output_path)
        shutil.move(video_output, output_path)

    return output_path, temp_files


# ═══════════════════════════════════════════════════════════════
# MODULE 3: VIDEO EDITOR
# ═══════════════════════════════════════════════════════════════

def build_trim_command(input_path, output_path, start, end):
    """Trim video from start to end (seconds)."""
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-ss', str(start),
        '-to', str(end),
        '-c', 'copy',  # try stream copy first for speed
        '-avoid_negative_ts', 'make_zero',
        output_path
    ]
    return cmd


def build_cut_command(input_path, output_path, segments):
    """
    Cut multiple segments and concatenate.
    segments: list of (start, end) tuples in seconds.
    """
    # Create temporary trimmed clips
    temp_dir = os.path.dirname(output_path)
    temp_files = []
    concat_list = os.path.join(temp_dir, f"cut_concat_{uuid.uuid4().hex}.txt")

    for idx, (start, end) in enumerate(segments):
        temp_path = os.path.join(temp_dir, f"segment_{idx}_{uuid.uuid4().hex}.mp4")
        cmd = build_trim_command(input_path, temp_path, start, end)
        run_ffmpeg(cmd, timeout=300)
        temp_files.append(temp_path)

    # Write concat list
    with open(concat_list, 'w') as f:
        for tf in temp_files:
            f.write(f"file '{os.path.abspath(tf)}'\n")

    # Concatenate
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-f', 'concat', '-safe', '0',
        '-i', concat_list,
        '-c', 'copy',
        output_path
    ]
    return cmd, temp_files + [concat_list]


def build_rotate_command(input_path, output_path, angle):
    """Rotate video by angle (90, 180, 270)."""
    transpose = {'90': '1', '180': '2,transpose=2', '270': '3'}
    t = transpose.get(str(angle), '1')
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-vf', f'transpose={t}',
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
        '-c:a', 'aac', '-b:a', '128k',
        '-pix_fmt', 'yuv420p',
        output_path
    ]
    return cmd


def build_resize_command(input_path, output_path, width, height):
    """Resize video to exact dimensions."""
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-vf', f'scale={width}:{height}',
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
        '-c:a', 'aac', '-b:a', '128k',
        '-pix_fmt', 'yuv420p',
        output_path
    ]
    return cmd


def build_crop_command(input_path, output_path, x, y, width, height):
    """Crop video region."""
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-vf', f'crop={width}:{height}:{x}:{y}',
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
        '-c:a', 'aac', '-b:a', '128k',
        '-pix_fmt', 'yuv420p',
        output_path
    ]
    return cmd


def build_speed_command(input_path, output_path, speed_factor):
    """Change video speed. speed_factor: 0.5 = half, 2.0 = double."""
    # atempo filter range is 0.5 to 100, chain for <0.5
    if speed_factor >= 0.5:
        atempo = f'atempo={1/speed_factor}'
    elif speed_factor >= 0.25:
        atempo = f'atempo=0.5,atempo={0.5/speed_factor}'
    else:
        atempo = f'atempo=0.5,atempo=0.5,atempo={0.25/speed_factor}'

    vf = f"setpts={1/speed_factor}*PTS"
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-vf', vf,
        '-af', atempo,
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
        '-c:a', 'aac', '-b:a', '128k',
        '-pix_fmt', 'yuv420p',
        output_path
    ]
    return cmd


def build_mute_command(input_path, output_path):
    """Remove audio from video."""
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-c:v', 'copy',
        '-an',  # no audio
        output_path
    ]
    return cmd


def build_replace_audio_command(input_path, audio_path, output_path):
    """Replace video audio track."""
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-i', audio_path,
        '-c:v', 'copy',
        '-map', '0:v:0',
        '-map', '1:a:0',
        '-c:a', 'aac', '-b:a', '192k',
        '-shortest',
        output_path
    ]
    return cmd


# ═══════════════════════════════════════════════════════════════
# MODULE 4: VIDEO COMPRESSOR
# ═══════════════════════════════════════════════════════════════

def build_compress_command(input_path, output_path, options=None):
    """Build FFmpeg command for video compression."""
    options = options or {}
    quality = options.get('quality', 'medium')  # low, medium, high
    codec = options.get('codec', 'libx264')     # libx264 or libx265
    target_size_mb = options.get('target_size_mb')

    preset = QUALITY_PRESETS.get(quality, QUALITY_PRESETS['medium'])

    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-threads', '0',
        '-c:v', codec,
        '-preset', preset['preset'],
        '-crf', str(preset['crf']),
        '-c:a', 'aac', '-b:a', '96k',
        '-pix_fmt', 'yuv420p',
        '-movflags', '+faststart',
    ]

    if target_size_mb:
        # Two-pass encoding for target size
        info = get_video_info(input_path)
        duration = info.get('duration', 0)
        if duration > 0:
            target_bitrate = (target_size_mb * 8192) / duration  # kbps roughly
            cmd = [
                '-y', '-hide_banner', '-loglevel', 'error',
                '-i', input_path,
                '-threads', '0',
                '-c:v', codec,
                '-preset', preset['preset'],
                '-b:v', f'{int(target_bitrate)}k',
                '-maxrate', f'{int(target_bitrate * 1.5)}k',
                '-bufsize', f'{int(target_bitrate)}k',
                '-pass', '1',
                '-an',
                '-f', 'null',
                '/dev/null' if os.name != 'nt' else 'NUL',
            ]
            # Second pass will be called separately

    cmd.append(output_path)
    return cmd


# ═══════════════════════════════════════════════════════════════
# MODULE 5: VIDEO MERGER
# ═══════════════════════════════════════════════════════════════

def build_merge_command(video_paths, output_path, options=None):
    """Merge multiple videos into one."""
    options = options or {}
    concat_file = os.path.join(os.path.dirname(output_path), f"merge_{uuid.uuid4().hex}.txt")

    with open(concat_file, 'w') as f:
        for vp in video_paths:
            f.write(f"file '{os.path.abspath(vp)}'\n")

    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-f', 'concat', '-safe', '0',
        '-i', concat_file,
        '-c', 'copy',
        '-movflags', '+faststart',
        output_path
    ]
    return cmd, concat_file


# ═══════════════════════════════════════════════════════════════
# MODULE 6: GIF MAKER
# ═══════════════════════════════════════════════════════════════

def build_gif_command(input_path, output_path, options=None):
    """Convert video to optimized GIF."""
    options = options or {}
    fps = options.get('fps', 10)
    width = options.get('width', 480)
    start = options.get('start', 0)
    duration = options.get('duration', 5)

    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-ss', str(start),
        '-t', str(duration),
        '-i', input_path,
        '-vf', f"fps={fps},scale={width}:-1:flags=lanczos,split[s0][s1];[s0]palettegen=max_colors=128[p];[s1][p]paletteuse=dither=bayer",
        '-loop', '0',
        output_path
    ]
    return cmd


# ═══════════════════════════════════════════════════════════════
# MODULE 7: AUDIO EXTRACTOR
# ═══════════════════════════════════════════════════════════════

def build_audio_extract_command(input_path, output_path, options=None):
    """Extract audio from video with advanced options."""
    options = options or {}
    format_ext = options.get('format', 'mp3')
    quality = options.get('quality', '192k')
    start = options.get('start', 0)
    duration = options.get('duration', 0)
    sample_rate = options.get('sample_rate')
    channels = options.get('channels')

    ext = format_ext.lower().lstrip('.')
    if ext == 'mp3':
        codec = 'libmp3lame'
    elif ext == 'wav':
        codec = 'pcm_s16le'
    elif ext == 'aac':
        codec = 'aac'
    elif ext == 'ogg':
        codec = 'libvorbis'
    elif ext == 'flac':
        codec = 'flac'
    elif ext == 'm4a':
        codec = 'aac'
    else:
        codec = 'libmp3lame'
        ext = 'mp3'

    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error'
    ]

    if start > 0:
        cmd += ['-ss', str(start)]
    if duration > 0:
        cmd += ['-t', str(duration)]

    cmd += ['-i', input_path, '-vn', '-c:a', codec]

    if codec == 'libmp3lame':
        cmd += ['-b:a', quality]
    elif codec == 'aac':
        cmd += ['-b:a', quality]
    elif codec == 'libvorbis':
        cmd += ['-q:a', '4']

    if sample_rate:
        cmd += ['-ar', str(sample_rate)]
    
    if channels == 'mono':
        cmd += ['-ac', '1']
    elif channels == 'stereo':
        cmd += ['-ac', '2']

    cmd.append(output_path)
    return cmd


# ═══════════════════════════════════════════════════════════════
# MODULE 8: WATERMARK
# ═══════════════════════════════════════════════════════════════

def build_watermark_command(input_path, output_path, options=None):
    """Add watermark to video."""
    options = options or {}
    text = options.get('text', 'Watermark')
    position = options.get('position', 'bottom-right')  # top-left, top-right, bottom-left, bottom-right, center
    font_size = options.get('font_size', 24)
    color = options.get('color', 'white')
    opacity = options.get('opacity', 0.7)
    image_path = options.get('image_path')

    # Build drawtext position
    pos_map = {
        'top-left':     f'x=10:y=10',
        'top-right':    f'x=W-tw-10:y=10',
        'bottom-left':  f'x=10:y=H-th-10',
        'bottom-right': f'x=W-tw-10:y=H-th-10',
        'center':       f'x=(W-tw)/2:y=(H-th)/2',
    }
    pos = pos_map.get(position, pos_map['bottom-right'])

    if image_path and os.path.exists(image_path):
        # Image watermark overlay
        # For image, we need to scale logo and overlay
        vf = f"movie={image_path}[wm];[in][wm]overlay={pos.replace('x=', '').replace('y=', '')}:format=auto"
        cmd = [
            '-y', '-hide_banner', '-loglevel', 'error',
            '-i', input_path,
            '-vf', vf,
            '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
            '-c:a', 'copy',
            output_path
        ]
    else:
        # Text watermark with drawtext (requires ffmpeg with libfreetype)
        # Fallback to subtitles/ass if drawtext unavailable
        vf = f"drawtext=text='{text}':{pos}:fontsize={font_size}:fontcolor={color}@{opacity}"
        cmd = [
            '-y', '-hide_banner', '-loglevel', 'error',
            '-i', input_path,
            '-vf', vf,
            '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
            '-c:a', 'copy',
            output_path
        ]
    return cmd


# ═══════════════════════════════════════════════════════════════
# MODULE 9: SUBTITLE OVERLAY
# ═══════════════════════════════════════════════════════════════

def build_subtitle_command(input_path, subtitle_path, output_path, options=None):
    """Burn subtitles into video."""
    options = options or {}
    font_size = options.get('font_size', 24)
    color = options.get('color', 'white')
    outline = options.get('outline', 1)
    bold = options.get('bold', 0)

    # Use subtitles filter for .srt/.ass, or ass filter for more control
    ext = os.path.splitext(subtitle_path)[1].lower()
    if ext in ('.srt', '.vtt'):
        vf = f"subtitles={subtitle_path}:force_style='FontSize={font_size},PrimaryColour=&H00{color},Outline={outline},Bold={bold}'"
    else:
        vf = f"ass={subtitle_path}"

    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-vf', vf,
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
        '-c:a', 'copy',
        output_path
    ]
    return cmd


# ═══════════════════════════════════════════════════════════════
# TEXT OVERLAY (for video editor)
# ═══════════════════════════════════════════════════════════════

def build_text_overlay_command(input_path, output_path, text, options=None):
    """Add text overlay to video."""
    options = options or {}
    position = options.get('position', 'center')
    font_size = options.get('font_size', 36)
    color = options.get('color', 'white')
    start_time = options.get('start', 0)
    duration = options.get('duration', 5)

    pos_map = {
        'top':          'x=(W-tw)/2:y=20',
        'bottom':       'x=(W-tw)/2:y=H-th-20',
        'center':       'x=(W-tw)/2:y=(H-th)/2',
        'top-left':     'x=20:y=20',
        'top-right':    'x=W-tw-20:y=20',
        'bottom-left':  'x=20:y=H-th-20',
        'bottom-right': 'x=W-tw-20:y=H-th-20',
    }
    pos = pos_map.get(position, pos_map['center'])

    vf = f"drawtext=text='{text}':{pos}:fontsize={font_size}:fontcolor={color}:enable='between(t,{start_time},{start_time+duration})'"
    cmd = [
        '-y', '-hide_banner', '-loglevel', 'error',
        '-i', input_path,
        '-vf', vf,
        '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
        '-c:a', 'copy',
        output_path
    ]
    return cmd
