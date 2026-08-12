import logging
import os

from django.conf import settings
from django.contrib import messages
from django.shortcuts import render

from .forms import TrimVideoForm
from .services.video.video_trim import trim_video


logger = logging.getLogger(__name__)




from .forms import MergeVideoForm
from .services.video.video_merge import merge_videos


logger = logging.getLogger(__name__)

def trim_video_view(request):
    """
    Handle video trimming requests.
    """

    if request.method == "POST":

        form = TrimVideoForm(
            request.POST,
            request.FILES,
        )

        if not form.is_valid():

            messages.error(
                request,
                "Please correct the errors and try again.",
            )

            return render(
                request,
                "media_tools/trim_video.html",
                {
                    "form": form,
                },
            )

        try:

            video = form.cleaned_data["video"]
            start_time = form.cleaned_data["start_time"]
            end_time = form.cleaned_data["end_time"]

            logger.info(
                "Trim request received: %s",
                video.name,
            )

            output_video = trim_video(
                video_file=video,
                start_time=start_time,
                end_time=end_time,
            )

            output_video_url = (
                settings.MEDIA_URL
                + output_video
                .replace(
                    settings.MEDIA_ROOT,
                    "",
                )
                .replace("\\", "/")
                .lstrip("/")
            )

            messages.success(
                request,
                "Video trimmed successfully.",
            )

            return render(
                request,
                "media_tools/trim_video_result.html",
                {
                    "video_url": output_video_url,
                },
            )

        except Exception as exc:

            logger.exception(
                "Video trimming request failed."
            )

            messages.error(
                request,
                str(exc),
            )

            return render(
                request,
                "media_tools/trim_video.html",
                {
                    "form": form,
                },
            )

    form = TrimVideoForm()

    return render(
        request,
        "media_tools/trim_video.html",
        {
            "form": form,
        },
    )



def merge_video_view(request):
    """
    Handle multiple video uploads and merge them into one video.
    """

    if request.method == "POST":

        form = MergeVideoForm(request.POST, request.FILES)

        # --------------------------------------------------
        # Get all uploaded videos
        # --------------------------------------------------

        video_files = request.FILES.getlist("videos")

        # --------------------------------------------------
        # Validate number of videos
        # --------------------------------------------------

        if len(video_files) < 2:

            messages.error(
                request,
                "Please upload at least two videos."
            )

            return render(
                request,
                "media_tools/merge_video.html",
                {
                    "form": form,
                }
            )

        # --------------------------------------------------
        # Validate form
        # --------------------------------------------------

        if not form.is_valid():

            messages.error(
                request,
                "Please check the uploaded video files."
            )

            return render(
                request,
                "media_tools/merge_video.html",
                {
                    "form": form,
                }
            )

        try:

            logger.info(
                "Video merge request received. Files: %d",
                len(video_files),
            )

            # --------------------------------------------------
            # Send videos to service layer
            # --------------------------------------------------

            output_video = merge_videos(
                video_files
            )

            # --------------------------------------------------
            # Convert absolute path to media URL
            # --------------------------------------------------

            media_root = str(
                request.META.get(
                    "MEDIA_ROOT",
                    ""
                )
            )

            # Use Django MEDIA_ROOT directly
            from django.conf import settings

            relative_path = os.path.relpath(
                output_video,
                settings.MEDIA_ROOT,
            )

            relative_path = relative_path.replace(
                "\\",
                "/",
            )

            video_url = (
                settings.MEDIA_URL.rstrip("/")
                + "/"
                + relative_path.lstrip("/")
            )

            logger.info(
                "Merged video created: %s",
                output_video,
            )

            logger.info(
                "Merged video URL: %s",
                video_url,
            )

            messages.success(
                request,
                "Videos merged successfully."
            )

            # --------------------------------------------------
            # Result page
            # --------------------------------------------------

            return render(
                request,
                "media_tools/merge_result.html",
                {
                    "video_url": video_url,
                }
            )

        except ValueError as exc:

            logger.warning(
                "Video merge validation/processing error: %s",
                exc,
            )

            messages.error(
                request,
                str(exc)
            )

            return render(
                request,
                "media_tools/merge_video.html",
                {
                    "form": form,
                }
            )

        except Exception:

            logger.exception(
                "Unexpected error in merge_video_view."
            )

            messages.error(
                request,
                "Unable to merge the videos. Please try again."
            )

            return render(
                request,
                "media_tools/merge_video.html",
                {
                    "form": form,
                }
            )

    # ------------------------------------------------------
    # GET request
    # ------------------------------------------------------

    form = MergeVideoForm()

    return render(
        request,
        "media_tools/merge_video.html",
        {
            "form": form,
        }
    )