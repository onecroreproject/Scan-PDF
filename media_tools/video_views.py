import logging
from pathlib import Path
from django.conf import settings
from django.contrib import messages
from django.http import Http404, FileResponse, JsonResponse
from django.shortcuts import render

from media_tools.services.file_service import get_video_directories
from media_tools.utils.card_helpers import get_tool_cards

from .video_forms import (
    CropVideoForm, ResizeVideoForm, TrimVideoForm, RotateVideoForm,
    FlipVideoForm, SpeedVideoForm, CompressVideoForm, ConvertVideoForm,
    VideoToGifForm, GifToVideoForm, RemoveAudioForm, ExtractAudioForm,
    ChangeVolumeForm, ReverseVideoForm, FadeVideoForm, EffectsVideoForm,
    FpsVideoForm, LoopVideoForm, FrameExtractForm, WatermarkVideoForm,
    MergeVideoForm, AddAudioForm, AddTextForm, AddImageForm, BlurVideoForm,
    BrightenVideoForm, SplitScreenForm, StabilizeVideoForm, SyncAudioForm,
    TransitionVideoForm, RepairVideoForm, BackgroundVideoForm,
)

from media_tools.services.video.crop_service import process_crop
from media_tools.services.video.resize_service import process_resize
from media_tools.services.video.trim import process_trim
from media_tools.services.video.rotate import process_rotate
from media_tools.services.video.flip import process_flip
from media_tools.services.video.speed import process_speed
from media_tools.services.video.compress import process_compress
from media_tools.services.video.convert import process_convert
from media_tools.services.video.fps_service import process_fps
from media_tools.services.video.audio_service import (
    process_remove_audio, process_extract_audio, process_change_volume,
)
from media_tools.services.video.gif_service import (
    process_video_to_gif, process_gif_to_video,
)
from media_tools.services.video.reverse_service import process_reverse
from media_tools.services.video.fade_service import process_fade
from media_tools.services.video.effects_service import process_effects
from media_tools.services.video.loop_service import process_loop
from media_tools.services.video.frame_service import process_extract_frame
from media_tools.services.video.watermark_service import process_watermark
from media_tools.services.video.video_merge import merge_videos
from media_tools.services.video.add_audio_service import process_add_audio
from media_tools.services.video.add_text_service import process_add_text
from media_tools.services.video.add_image_service import process_add_image
from media_tools.services.video.blur_service import process_blur
from media_tools.services.video.brighten_service import process_brighten
from media_tools.services.video.split_screen_service import process_split_screen
from media_tools.services.video.stabilize_service import process_stabilize
from media_tools.services.video.sync_service import process_sync
from media_tools.services.video.transition_service import process_transition
from media_tools.services.video.repair_service import process_repair
from media_tools.services.video.background_service import process_background

logger = logging.getLogger("media_tools")


def _handle_tool_view(request, form_class, runner_fn, template_name, extra_ctx=None):
    extra = extra_ctx or {}
    slug = extra.get("tool_slug", "")
    cards = extra.get("cards") or (get_tool_cards(slug) if slug else [])
    ctx = {"cards": cards, **extra}

    if request.method == "GET":
        return render(request, template_name, {"form": form_class(), **ctx})

    is_ajax = request.headers.get("X-Requested-With") == "XMLHttpRequest"
    form = form_class(request.POST, request.FILES)

    if not form.is_valid():
        first_err = next(iter(form.errors.values()))[0] if form.errors else "Please check your inputs."
        if is_ajax:
            return JsonResponse({"success": False, "message": str(first_err), "errors": form.errors.get_json_data()}, status=400)
        messages.error(request, str(first_err))
        return render(request, template_name, {"form": form, **ctx}, status=400)

    try:
        output_path = runner_fn(form.cleaned_data)
        if not output_path:
            raise RuntimeError("Media processing did not produce an output file.")

        filename = output_path.name if hasattr(output_path, "name") else Path(output_path).name
        output_url = f"/media/video_tools/outputs/{filename}"
        download_url = f"/media_tools/download/{filename}/"

        if is_ajax:
            return JsonResponse({
                "success": True,
                "message": "Processed successfully.",
                "video_url": output_url,
                "download_url": output_url,
                "filename": filename,
            })

        messages.success(request, "Processed successfully.")
        return render(request, template_name, {
            "form": form_class(),
            "output_url": output_url,
            "download_url": output_url,
            "filename": filename,
            **ctx,
        })

    except ValueError as exc:
        logger.warning("Validation error in %s: %s", template_name, exc)
        if is_ajax:
            return JsonResponse({"success": False, "message": str(exc)}, status=400)
        messages.error(request, str(exc))
        return render(request, template_name, {"form": form, **ctx}, status=400)

    except Exception as exc:
        logger.exception("Processing error in %s: %s", template_name, exc)
        msg = str(exc) if isinstance(exc, ValueError) else "Unable to process this file. Please check settings and try again."
        if is_ajax:
            return JsonResponse({"success": False, "message": msg}, status=500)
        messages.error(request, msg)
        return render(request, template_name, {"form": form, **ctx}, status=500)


# =============================================================================
# 32 VIDEO & MEDIA TOOL VIEWS
# =============================================================================

def crop_video(request):
    """Crop video view."""
    return _handle_tool_view(
        request, CropVideoForm,
        lambda d: process_crop(
            video=d["video"], x=d["x"], y=d["y"],
            width=d["width"], height=d["height"],
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/crop.html",
        {"tool_name": "Crop Video", "tool_slug": "crop"},
    )


def resize_video(request):
    """Resize video view."""
    return _handle_tool_view(
        request, ResizeVideoForm,
        lambda d: process_resize(
            video=d["video"], width=d["width"], height=d["height"],
            aspect_ratio=d.get("aspect_ratio", ""), fit_mode=d.get("fit_mode", "fit"),
            zoom=d.get("zoom", 1.0), position_x=d.get("position_x", 0),
            position_y=d.get("position_y", 0), background_color=d.get("background_color", "#000000"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/resize.html",
        {"tool_name": "Resize Video", "tool_slug": "resize"},
    )


def trim_video(request):
    """Trim video view."""
    return _handle_tool_view(
        request, TrimVideoForm,
        lambda d: process_trim(
            video=d["video"], start_time=d.get("start_time"),
            end_time=d.get("end_time"), output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/trim.html",
        {"tool_name": "Trim Video", "tool_slug": "trim"},
    )


def rotate_video(request):
    """Rotate video view."""
    return _handle_tool_view(
        request, RotateVideoForm,
        lambda d: process_rotate(
            video=d["video"], rotation=d.get("rotation", "90"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/rotate.html",
        {"tool_name": "Rotate Video", "tool_slug": "rotate"},
    )


def flip_video(request):
    """Flip video horizontally or vertically."""
    return _handle_tool_view(
        request, FlipVideoForm,
        lambda d: process_flip(
            d["video"], flip_mode=d.get("flip_mode", "horizontal"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/flip.html",
        {"tool_name": "Flip Video", "tool_slug": "flip"},
    )


def speed_video(request):
    """Adjust video playback speed (slow-mo / speed-up)."""
    return _handle_tool_view(
        request, SpeedVideoForm,
        lambda d: process_speed(
            d["video"], speed=d.get("speed", 1.0),
            keep_audio=d.get("keep_audio", True),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/speed.html",
        {"tool_name": "Change Video Speed", "tool_slug": "speed"},
    )


def compress_video(request):
    """Compress video file size."""
    return _handle_tool_view(
        request, CompressVideoForm,
        lambda d: process_compress(
            d["video"], compression_level=d.get("compression_level", "medium"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/compress.html",
        {"tool_name": "Compress Video", "tool_slug": "compress"},
    )


def convert_video(request):
    """Convert video format."""
    return _handle_tool_view(
        request, ConvertVideoForm,
        lambda d: process_convert(
            d["video"], target_format=d.get("target_format", "mp4"),
        ),
        "media_tools/convert.html",
        {"tool_name": "Convert Video Format", "tool_slug": "convert"},
    )


def video_to_gif(request):
    """Convert video clip to animated GIF."""
    return _handle_tool_view(
        request, VideoToGifForm,
        lambda d: process_video_to_gif(
            d["video"], fps=d.get("fps", 15), width=d.get("width", 480),
        ),
        "media_tools/video_to_gif.html",
        {"tool_name": "Video to GIF", "tool_slug": "video-to-gif"},
    )


def gif_to_video(request):
    """Convert animated GIF to MP4 video."""
    return _handle_tool_view(
        request, GifToVideoForm,
        lambda d: process_gif_to_video(
            d["video"], output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/gif_to_video.html",
        {"tool_name": "GIF to Video", "tool_slug": "gif-to-video"},
    )


def remove_audio(request):
    """Mute / remove audio tracks from video."""
    return _handle_tool_view(
        request, RemoveAudioForm,
        lambda d: process_remove_audio(
            d["video"], output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/remove_audio.html",
        {"tool_name": "Remove Audio", "tool_slug": "remove-audio"},
    )


def extract_audio(request):
    """Extract audio soundtrack from video."""
    return _handle_tool_view(
        request, ExtractAudioForm,
        lambda d: process_extract_audio(
            d["video"], audio_format=d.get("audio_format", "mp3"),
            bitrate=d.get("audio_bitrate", "192k"),
        ),
        "media_tools/extract_audio.html",
        {"tool_name": "Extract Audio", "tool_slug": "extract-audio"},
    )


def change_volume(request):
    """Adjust video volume."""
    return _handle_tool_view(
        request, ChangeVolumeForm,
        lambda d: process_change_volume(
            d["video"], volume=d.get("volume", 100),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/change_volume.html",
        {"tool_name": "Change Volume", "tool_slug": "change-volume"},
    )


def reverse_video(request):
    """Reverse video playback."""
    return _handle_tool_view(
        request, ReverseVideoForm,
        lambda d: process_reverse(
            d["video"], reverse_audio=d.get("reverse_audio", True),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/reverse.html",
        {"tool_name": "Reverse Video", "tool_slug": "reverse"},
    )


def fade_video(request):
    """Apply fade-in and fade-out to video."""
    return _handle_tool_view(
        request, FadeVideoForm,
        lambda d: process_fade(
            d["video"], fade_in_duration=d.get("fade_in", 1.0),
            fade_out_duration=d.get("fade_out", 1.0),
            color=d.get("fade_color", "black"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/fade.html",
        {"tool_name": "Video Fade In/Out", "tool_slug": "fade"},
    )


def effects_video(request):
    """Adjust brightness, contrast, saturation, or artistic filters."""
    return _handle_tool_view(
        request, EffectsVideoForm,
        lambda d: process_effects(
            d["video"], brightness=d.get("brightness", 0.0),
            contrast=d.get("contrast", 1.0), saturation=d.get("saturation", 1.0),
            filter_type=d.get("filter_type", "none"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/effects.html",
        {"tool_name": "Video Filters & Effects", "tool_slug": "effects"},
    )


def fps_video(request):
    """Change video frame rate (FPS)."""
    return _handle_tool_view(
        request, FpsVideoForm,
        lambda d: process_fps(
            d["video"], fps=d.get("fps", 30),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/fps.html",
        {"tool_name": "Change Frame Rate (FPS)", "tool_slug": "fps"},
    )


def loop_video(request):
    """Loop video N times."""
    return _handle_tool_view(
        request, LoopVideoForm,
        lambda d: process_loop(
            d["video"], loop_count=d.get("loop_count", 2),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/loop.html",
        {"tool_name": "Loop Video", "tool_slug": "loop"},
    )


def extract_frame(request):
    """Extract pristine still frame or screenshot from video."""
    return _handle_tool_view(
        request, FrameExtractForm,
        lambda d: process_extract_frame(
            d["video"], time_offset=d.get("time_offset", 0.0),
            image_format=d.get("image_format", "jpg"),
        ),
        "media_tools/extract_frame.html",
        {"tool_name": "Video Screenshot / Frame", "tool_slug": "extract-frame"},
    )


def watermark_video(request):
    """Add watermark or logo overlay to video."""
    return _handle_tool_view(
        request, WatermarkVideoForm,
        lambda d: process_watermark(
            d["video"], watermark_image=d["watermark_image"],
            position=d.get("position", "bottom-right"),
            opacity=d.get("opacity", 0.8), scale_percent=d.get("scale", 15),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/watermark.html",
        {"tool_name": "Add Watermark", "tool_slug": "watermark"},
    )


def merge_video(request):
    """Merge multiple video clips into a single video."""
    cards = get_tool_cards("merge")
    if request.method == "GET":
        return render(
            request, "media_tools/merge_video.html",
            {"form": MergeVideoForm(), "tool_name": "Merge Video", "tool_slug": "merge", "cards": cards},
        )

    is_ajax = request.headers.get("X-Requested-With") == "XMLHttpRequest"
    video_files = request.FILES.getlist("videos") or ([request.FILES.get("video")] if request.FILES.get("video") else [])

    if len(video_files) < 2:
        msg = "Please upload at least 2 videos to merge."
        if is_ajax:
            return JsonResponse({"success": False, "message": msg}, status=400)
        messages.error(request, msg)
        return render(request, "media_tools/merge_video.html", {"form": MergeVideoForm(), "cards": cards}, status=400)

    try:
        output_path = merge_videos(video_files)
        filename = output_path.name if hasattr(output_path, "name") else Path(output_path).name
        output_url = f"/media/video_tools/outputs/{filename}"
        download_url = f"/media_tools/download/{filename}/"

        if is_ajax:
            return JsonResponse({
                "success": True,
                "message": "Videos merged successfully.",
                "video_url": output_url,
                "download_url": output_url,
                "filename": filename,
            })

        messages.success(request, "Videos merged successfully.")
        return render(request, "media_tools/merge_video.html", {
            "form": MergeVideoForm(),
            "output_url": output_url,
            "download_url": output_url,
            "filename": filename,
            "tool_name": "Merge Video",
            "tool_slug": "merge",
            "cards": cards,
        })
    except Exception as exc:
        logger.exception("Merge video failed: %s", exc)
        msg = str(exc) if isinstance(exc, ValueError) else "Unable to merge videos."
        if is_ajax:
            return JsonResponse({"success": False, "message": msg}, status=400 if isinstance(exc, ValueError) else 500)
        messages.error(request, msg)
        return render(request, "media_tools/merge_video.html", {"form": MergeVideoForm(), "cards": cards}, status=500)


def add_audio(request):
    """Add background audio or replacement soundtrack to video."""
    return _handle_tool_view(
        request, AddAudioForm,
        lambda d: process_add_audio(
            d["video"], audio=d["audio"], mode=d.get("mode", "replace"),
            audio_volume=d.get("audio_volume", 1.0),
            video_volume=d.get("video_volume", 1.0),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/add_audio.html",
        {"tool_name": "Add Audio", "tool_slug": "add-audio"},
    )


def add_text(request):
    """Add customized text overlay onto video."""
    return _handle_tool_view(
        request, AddTextForm,
        lambda d: process_add_text(
            d["video"], text=d.get("text", "Sample Text"),
            font_size=d.get("font_size", 36),
            font_color=d.get("font_color", "#ffffff"),
            position=d.get("position", "bottom-center"),
            bg_box=d.get("bg_box", False),
            x_percent=d.get("x_percent", 50),
            y_percent=d.get("y_percent", 85),
            start_time=d.get("start_time"),
            end_time=d.get("end_time"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/add_text.html",
        {"tool_name": "Add Text Overlay", "tool_slug": "add-text"},
    )


def add_image(request):
    """Add image watermark or logo overlay to video."""
    return _handle_tool_view(
        request, AddImageForm,
        lambda d: process_add_image(
            d["video"], image=d["image"],
            position=d.get("position", "bottom-right"),
            opacity=d.get("opacity", 0.85),
            scale_percent=d.get("scale_percent", 20),
            x_percent=d.get("x_percent", 80),
            y_percent=d.get("y_percent", 80),
            start_time=d.get("start_time"),
            end_time=d.get("end_time"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/add_image.html",
        {"tool_name": "Add Image / Logo", "tool_slug": "add-image"},
    )


def blur_video(request):
    """Blur whole video or rectangular area."""
    return _handle_tool_view(
        request, BlurVideoForm,
        lambda d: process_blur(
            d["video"], blur_type=d.get("blur_type", "full"),
            intensity=d.get("intensity", 15),
            x=d.get("x", 0), y=d.get("y", 0),
            width=d.get("width", 0), height=d.get("height", 0),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/blur.html",
        {"tool_name": "Blur Video", "tool_slug": "blur"},
    )


def brighten_video(request):
    """Adjust video brightness, contrast, saturation, and gamma."""
    return _handle_tool_view(
        request, BrightenVideoForm,
        lambda d: process_brighten(
            d["video"], brightness=d.get("brightness", 0.15),
            contrast=d.get("contrast", 1.0), saturation=d.get("saturation", 1.0),
            gamma=d.get("gamma", 1.0), output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/brighten.html",
        {"tool_name": "Brighten & Color Adjust", "tool_slug": "brighten"},
    )


def split_screen(request):
    """Combine videos into side-by-side or grid split screen."""
    return _handle_tool_view(
        request, SplitScreenForm,
        lambda d: process_split_screen(
            d["video1"], d["video2"], video3=d.get("video3"),
            video4=d.get("video4"), layout=d.get("layout", "2-side"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/split_screen.html",
        {"tool_name": "Split Screen Video", "tool_slug": "split-screen"},
    )


def stabilize_video(request):
    """Stabilize shaky camera movement."""
    return _handle_tool_view(
        request, StabilizeVideoForm,
        lambda d: process_stabilize(
            d["video"], strength=d.get("strength", "medium"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/stabilize.html",
        {"tool_name": "Stabilize Video", "tool_slug": "stabilize"},
    )


def sync_audio(request):
    """Synchronize out-of-sync audio and video tracks."""
    return _handle_tool_view(
        request, SyncAudioForm,
        lambda d: process_sync(
            d["video"], audio_offset=d.get("audio_offset", 0.0),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/sync_audio.html",
        {"tool_name": "Sync Audio & Video", "tool_slug": "sync-audio"},
    )


def transition_video(request):
    """Apply smooth video fade transitions."""
    return _handle_tool_view(
        request, TransitionVideoForm,
        lambda d: process_transition(
            d["video"], transition_type=d.get("transition_type", "fade-both"),
            duration=d.get("duration", 1.0), color=d.get("color", "black"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/transition.html",
        {"tool_name": "Add Video Transition", "tool_slug": "transition"},
    )


def repair_video(request):
    """Repair damaged, corrupt, or unplayable video files."""
    return _handle_tool_view(
        request, RepairVideoForm,
        lambda d: process_repair(
            d["video"], repair_mode=d.get("repair_mode", "smart"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/repair.html",
        {"tool_name": "Repair Corrupted Video", "tool_slug": "repair"},
    )


def background_video(request):
    """Apply blurred or solid color background padding for aspect ratios."""
    return _handle_tool_view(
        request, BackgroundVideoForm,
        lambda d: process_background(
            d["video"], bg_type=d.get("bg_type", "blur"),
            bg_color=d.get("bg_color", "#000000"),
            aspect_ratio=d.get("aspect_ratio", "9:16"),
            output_format=d.get("output_format", "mp4"),
        ),
        "media_tools/background.html",
        {"tool_name": "Video Background & Aspect Ratio", "tool_slug": "background"},
    )


def video_tool_view(request, tool_name):
    """Generic view that dispatches to the appropriate tool based on tool_name."""
    TOOL_MAP = {
        "crop": (CropVideoForm, lambda d: process_crop(video=d["video"], x=d["x"], y=d["y"], width=d["width"], height=d["height"], output_format=d.get("output_format", "mp4")), "media_tools/crop.html", {"tool_name": "Crop Video", "tool_slug": "crop"}),
        "resize": (ResizeVideoForm, lambda d: process_resize(video=d["video"], width=d["width"], height=d["height"], aspect_ratio=d.get("aspect_ratio", ""), fit_mode=d.get("fit_mode", "fit"), zoom=d.get("zoom", 1.0), position_x=d.get("position_x", 0), position_y=d.get("position_y", 0), background_color=d.get("background_color", "#000000"), output_format=d.get("output_format", "mp4")), "media_tools/resize.html", {"tool_name": "Resize Video", "tool_slug": "resize"}),
        "rotate": (RotateVideoForm, lambda d: process_rotate(video=d["video"], rotation=d.get("rotation", "90"), output_format=d.get("output_format", "mp4")), "media_tools/rotate.html", {"tool_name": "Rotate Video", "tool_slug": "rotate"}),
        "trim": (TrimVideoForm, lambda d: process_trim(video=d["video"], start_time=d.get("start_time"), end_time=d.get("end_time"), output_format=d.get("output_format", "mp4")), "media_tools/trim.html", {"tool_name": "Trim Video", "tool_slug": "trim"}),
        "add-audio": (AddAudioForm, lambda d: process_add_audio(d["video"], audio=d["audio"], mode=d.get("mode", "replace"), audio_volume=d.get("audio_volume", 1.0), video_volume=d.get("video_volume", 1.0), output_format=d.get("output_format", "mp4")), "media_tools/add_audio.html", {"tool_name": "Add Audio", "tool_slug": "add-audio"}),
        "add-text": (AddTextForm, lambda d: process_add_text(d["video"], text=d.get("text", "Sample Text"), font_size=d.get("font_size", 36), font_color=d.get("font_color", "#ffffff"), position=d.get("position", "bottom-center"), bg_box=d.get("bg_box", False), x_percent=d.get("x_percent", 50), y_percent=d.get("y_percent", 85), start_time=d.get("start_time"), end_time=d.get("end_time"), output_format=d.get("output_format", "mp4")), "media_tools/add_text.html", {"tool_name": "Add Text Overlay", "tool_slug": "add-text"}),
        "add-image": (AddImageForm, lambda d: process_add_image(d["video"], image=d["image"], position=d.get("position", "bottom-right"), opacity=d.get("opacity", 0.85), scale_percent=d.get("scale_percent", 20), x_percent=d.get("x_percent", 80), y_percent=d.get("y_percent", 80), start_time=d.get("start_time"), end_time=d.get("end_time"), output_format=d.get("output_format", "mp4")), "media_tools/add_image.html", {"tool_name": "Add Image / Logo", "tool_slug": "add-image"}),
        "remove-audio": (RemoveAudioForm, lambda d: process_remove_audio(d["video"], output_format=d.get("output_format", "mp4")), "media_tools/remove_audio.html", {"tool_name": "Remove Audio", "tool_slug": "remove-audio"}),
        "extract-audio": (ExtractAudioForm, lambda d: process_extract_audio(d["video"], audio_format=d.get("audio_format", "mp3"), bitrate=d.get("audio_bitrate", "192k")), "media_tools/extract_audio.html", {"tool_name": "Extract Audio", "tool_slug": "extract-audio"}),
        "change-volume": (ChangeVolumeForm, lambda d: process_change_volume(d["video"], volume=d.get("volume", 100), output_format=d.get("output_format", "mp4")), "media_tools/change_volume.html", {"tool_name": "Change Volume", "tool_slug": "change-volume"}),
    }
    clean_slug = (tool_name or "").strip().lower().replace("_", "-")
    if clean_slug not in TOOL_MAP:
        raise Http404(f"Tool '{tool_name}' not found.")
    form_class, runner_fn, template_name, extra_ctx = TOOL_MAP[clean_slug]
    return _handle_tool_view(request, form_class, runner_fn, template_name, extra_ctx)


def download_output(request, filename):
    """Serve processed media files safely as attachments."""
    _, outputs_dir, _ = get_video_directories()
    file_path = outputs_dir / filename
    if not file_path.exists():
        raise Http404("Requested file does not exist.")
    import mimetypes
    mime_type, _ = mimetypes.guess_type(str(file_path))
    mime_type = mime_type or "application/octet-stream"
    return FileResponse(open(file_path, "rb"), as_attachment=True, filename=filename, content_type=mime_type)
