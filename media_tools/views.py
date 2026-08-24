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


# media_tools/views.py

import logging
import os
import uuid

from django.conf import settings
from django.contrib import messages
from django.shortcuts import render, redirect

from .forms import CropVideoForm
from .services.video.video_crop import crop_video


logger = logging.getLogger(__name__)


def crop_video_view(request):
    """
    Video Crop workflow:

    GET:
        Display crop form and video preview UI.

    POST:
        1. Validate uploaded video.
        2. Validate crop settings.
        3. Send the request to video_crop service.
        4. Redirect to result page.
    """

    if request.method == "GET":
        form = CropVideoForm()

        return render(
            request,
            "media_tools/crop_video.html",
            {
                "form": form,
            },
        )

    # ---------------------------------------------------------
    # POST
    # ---------------------------------------------------------
    form = CropVideoForm(request.POST, request.FILES)

    if not form.is_valid():
        logger.warning(
            "Video crop form validation failed: %s",
            form.errors.as_json(),
        )

        return render(
            request,
            "media_tools/crop_video.html",
            {
                "form": form,
            },
            status=400,
        )

    try:
        # -----------------------------------------------------
        # Get validated form data
        # -----------------------------------------------------
        video = form.cleaned_data["video"]

        ratio = form.cleaned_data.get("ratio") or "free"

        custom_crop = form.cleaned_data.get(
            "custom_crop",
            False,
        )

        custom_width = form.cleaned_data.get(
            "custom_width"
        )

        custom_height = form.cleaned_data.get(
            "custom_height"
        )

        crop_x = form.cleaned_data.get("crop_x")
        crop_y = form.cleaned_data.get("crop_y")
        crop_width = form.cleaned_data.get("crop_width")
        crop_height = form.cleaned_data.get("crop_height")

        logger.info(
            "Video crop request received. "
            "filename=%s ratio=%s custom_crop=%s "
            "crop=(%s,%s,%s,%s)",
            getattr(video, "name", "unknown"),
            ratio,
            custom_crop,
            crop_x,
            crop_y,
            crop_width,
            crop_height,
        )

        # -----------------------------------------------------
        # Validate custom crop dimensions
        # -----------------------------------------------------
        if custom_crop:

            if not custom_width or not custom_height:
                form.add_error(
                    None,
                    "Please enter both custom crop width and height.",
                )

                logger.warning(
                    "Custom crop dimensions missing."
                )

                return render(
                    request,
                    "media_tools/crop_video.html",
                    {
                        "form": form,
                    },
                    status=400,
                )

        # -----------------------------------------------------
        # Validate crop rectangle
        # -----------------------------------------------------
        if not custom_crop:

            if (
                crop_x is None
                or crop_y is None
                or crop_width is None
                or crop_height is None
            ):
                form.add_error(
                    None,
                    "Please select a crop area on the video.",
                )

                logger.warning(
                    "Crop coordinates missing."
                )

                return render(
                    request,
                    "media_tools/crop_video.html",
                    {
                        "form": form,
                    },
                    status=400,
                )

            if crop_width <= 0 or crop_height <= 0:
                form.add_error(
                    None,
                    "Crop width and height must be greater than zero.",
                )

                logger.warning(
                    "Invalid crop dimensions: width=%s height=%s",
                    crop_width,
                    crop_height,
                )

                return render(
                    request,
                    "media_tools/crop_video.html",
                    {
                        "form": form,
                    },
                    status=400,
                )

            if crop_x < 0 or crop_y < 0:
                form.add_error(
                    None,
                    "Crop position cannot be negative.",
                )

                logger.warning(
                    "Invalid crop position: x=%s y=%s",
                    crop_x,
                    crop_y,
                )

                return render(
                    request,
                    "media_tools/crop_video.html",
                    {
                        "form": form,
                    },
                    status=400,
                )

        # -----------------------------------------------------
        # Create temporary upload directory
        # -----------------------------------------------------
        upload_dir = os.path.join(
            settings.MEDIA_ROOT,
            "media_tools",
            "crop_uploads",
        )

        os.makedirs(
            upload_dir,
            exist_ok=True,
        )

        # -----------------------------------------------------
        # Save uploaded video
        # -----------------------------------------------------
        unique_name = (
            f"{uuid.uuid4().hex}_"
            f"{os.path.basename(video.name)}"
        )

        input_path = os.path.join(
            upload_dir,
            unique_name,
        )

        with open(input_path, "wb+") as destination:

            for chunk in video.chunks():
                destination.write(chunk)

        logger.info(
            "Crop input video saved: %s",
            input_path,
        )

        # -----------------------------------------------------
        # Call video processing service
        # -----------------------------------------------------
        result = crop_video(
            input_path=input_path,
            ratio=ratio,
            custom_crop=custom_crop,
            custom_width=custom_width,
            custom_height=custom_height,
            crop_x=crop_x,
            crop_y=crop_y,
            crop_width=crop_width,
            crop_height=crop_height,
        )

        logger.info(
            "Video crop completed successfully: %s",
            result,
        )

        # -----------------------------------------------------
        # Result handling
        # -----------------------------------------------------
        if isinstance(result, dict):

            result_path = result.get("output_path")
            result_url = result.get("output_url")

        else:

            result_path = result
            result_url = None

        if not result_path:
            raise ValueError(
                "Crop service did not return an output video."
            )

        # -----------------------------------------------------
        # Store result in session
        # -----------------------------------------------------
        request.session["crop_result"] = {
            "output_path": result_path,
            "output_url": result_url,
            "original_filename": video.name,
            "ratio": ratio,
        }

        request.session.modified = True

        logger.info(
            "Redirecting user to crop result page."
        )

        return redirect(
            "media_tools:crop_result"
        )

    # ---------------------------------------------------------
    # Expected processing errors
    # ---------------------------------------------------------
    except ValueError as exc:

        logger.warning(
            "Video crop validation/processing error: %s",
            exc,
            exc_info=True,
        )

        form.add_error(
            None,
            str(exc),
        )

        return render(
            request,
            "media_tools/crop_video.html",
            {
                "form": form,
            },
            status=400,
        )

    # ---------------------------------------------------------
    # Unexpected errors
    # ---------------------------------------------------------
    except Exception as exc:

        logger.exception(
            "Unexpected video crop error."
        )

        form.add_error(
            None,
            "The video could not be cropped. "
            "Please try another video.",
        )

        return render(
            request,
            "media_tools/crop_video.html",
            {
                "form": form,
            },
            status=500,
        )