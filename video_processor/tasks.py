"""
Celery background tasks for video processing.
Handles heavy FFmpeg operations asynchronously.
"""
import os
import logging
from celery import shared_task
from django.conf import settings

from .ffmpeg_helpers import (
    run_ffmpeg,
    build_convert_command,
    build_image_to_video_command,
    build_compress_command,
    build_merge_command,
    build_gif_command,
    build_audio_extract_command,
    build_watermark_command,
    build_subtitle_command,
    build_text_overlay_command,
    build_trim_command,
    build_rotate_command,
    build_resize_command,
    build_crop_command,
    build_speed_command,
    build_mute_command,
    build_replace_audio_command,
    build_cut_command,
)

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════
# TASK HELPERS
# ═══════════════════════════════════════════════════════════════

def _safe_remove(path):
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass

# ═══════════════════════════════════════════════════════════════
# MODULE 1: VIDEO CONVERTER TASK
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def convert_video_task(self, input_path, output_path, output_format, options=None):
    """Async video conversion."""
    try:
        cmd = build_convert_command(input_path, output_path, output_format, options)
        run_ffmpeg(cmd, timeout=1800)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


# ═══════════════════════════════════════════════════════════════
# MODULE 2: IMAGE TO VIDEO TASK
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def image_to_video_task(self, image_paths, output_path, options=None):
    """Async image to video creation."""
    concat_file = None
    try:
        cmd, concat_file = build_image_to_video_command(image_paths, output_path, options)
        run_ffmpeg(cmd, timeout=1800)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        if concat_file:
            _safe_remove(concat_file)
        for ip in image_paths:
            _safe_remove(ip)


# ═══════════════════════════════════════════════════════════════
# MODULE 3: VIDEO EDITOR TASKS
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def trim_video_task(self, input_path, output_path, start, end):
    try:
        cmd = build_trim_command(input_path, output_path, start, end)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def cut_video_task(self, input_path, output_path, segments):
    temp_files = []
    try:
        cmd, temp_files = build_cut_command(input_path, output_path, segments)
        run_ffmpeg(cmd, timeout=900)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        for tf in temp_files:
            _safe_remove(tf)
        _safe_remove(input_path)


@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def rotate_video_task(self, input_path, output_path, angle):
    try:
        cmd = build_rotate_command(input_path, output_path, angle)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def resize_video_task(self, input_path, output_path, width, height):
    try:
        cmd = build_resize_command(input_path, output_path, width, height)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def crop_video_task(self, input_path, output_path, x, y, width, height):
    try:
        cmd = build_crop_command(input_path, output_path, x, y, width, height)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def speed_video_task(self, input_path, output_path, speed_factor):
    try:
        cmd = build_speed_command(input_path, output_path, speed_factor)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def mute_video_task(self, input_path, output_path):
    try:
        cmd = build_mute_command(input_path, output_path)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def replace_audio_task(self, input_path, audio_path, output_path):
    try:
        cmd = build_replace_audio_command(input_path, audio_path, output_path)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)
        _safe_remove(audio_path)


@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def text_overlay_task(self, input_path, output_path, text, options=None):
    try:
        cmd = build_text_overlay_command(input_path, output_path, text, options)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


# ═══════════════════════════════════════════════════════════════
# MODULE 4: VIDEO COMPRESSOR TASK
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def compress_video_task(self, input_path, output_path, options=None):
    try:
        cmd = build_compress_command(input_path, output_path, options)
        run_ffmpeg(cmd, timeout=1800)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


# ═══════════════════════════════════════════════════════════════
# MODULE 5: VIDEO MERGER TASK
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def merge_videos_task(self, video_paths, output_path, options=None):
    concat_file = None
    try:
        cmd, concat_file = build_merge_command(video_paths, output_path, options)
        run_ffmpeg(cmd, timeout=1800)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        if concat_file:
            _safe_remove(concat_file)
        for vp in video_paths:
            _safe_remove(vp)


# ═══════════════════════════════════════════════════════════════
# MODULE 6: GIF MAKER TASK
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def make_gif_task(self, input_path, output_path, options=None):
    try:
        cmd = build_gif_command(input_path, output_path, options)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


# ═══════════════════════════════════════════════════════════════
# MODULE 7: AUDIO EXTRACTOR TASK
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def extract_audio_task(self, input_path, output_path, format_ext, quality='192k'):
    try:
        cmd = build_audio_extract_command(input_path, output_path, format_ext, quality)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)


# ═══════════════════════════════════════════════════════════════
# MODULE 8: WATERMARK TASK
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def add_watermark_task(self, input_path, output_path, options=None):
    try:
        cmd = build_watermark_command(input_path, output_path, options)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)
        if options and options.get('image_path'):
            _safe_remove(options.get('image_path'))


# ═══════════════════════════════════════════════════════════════
# MODULE 9: SUBTITLE OVERLAY TASK
# ═══════════════════════════════════════════════════════════════

@shared_task(bind=True, max_retries=2, default_retry_delay=30)
def add_subtitle_task(self, input_path, subtitle_path, output_path, options=None):
    try:
        cmd = build_subtitle_command(input_path, subtitle_path, output_path, options)
        run_ffmpeg(cmd, timeout=600)
        return {'status': 'success', 'output_path': output_path}
    except Exception as exc:
        _safe_remove(output_path)
        raise self.retry(exc=exc)
    finally:
        _safe_remove(input_path)
        _safe_remove(subtitle_path)
