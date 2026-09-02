
import logging

from django.conf import settings
from django.contrib import messages
from django.shortcuts import render

from .video_forms import CropVideoForm
from .video_forms import (
    ResizeVideoForm,
)

from media_tools.services.video.crop_service import (
    process_crop,
)
from media_tools.services.video.resize_service import (
    process_resize,
)


logger = logging.getLogger("media_tools")



def crop_video(request):
    """
    Crop uploaded video.
    """

    if request.method == "POST":

        form = CropVideoForm(
            request.POST,
            request.FILES,
        )

        if form.is_valid():

            try:

                output_path = process_crop(
                    video=form.cleaned_data["video"],
                    x=form.cleaned_data["x"],
                    y=form.cleaned_data["y"],
                    width=form.cleaned_data["width"],
                    height=form.cleaned_data["height"],
                )

                output_url = (
                    settings.MEDIA_URL
                    + "video_tools/outputs/"
                    + output_path.name
                )

                messages.success(
                    request,
                    "Video cropped successfully.",
                )

                return render(
                    request,
                    "media_tools/crop.html",
                    {
                        "form": CropVideoForm(),
                        "output_url": output_url,
                    },
                )

            except ValueError as exc:

                logger.warning(
                    "Crop validation error: %s",
                    exc,
                )

                form.add_error(
                    None,
                    str(exc),
                )

                messages.error(
                    request,
                    str(exc),
                )

            except Exception:

                logger.exception(
                    "Unexpected crop view error."
                )

                form.add_error(
                    None,
                    "Unable to process the video.",
                )

                messages.error(
                    request,
                    "Unable to process the video.",
                )

    else:

        form = CropVideoForm()

    return render(
        request,
        "media_tools/crop.html",
        {
            "form": form,
        },
    )

def resize_video(request):
    """
    Resize uploaded video using editor settings.
    """

    if request.method == "POST":

        form = ResizeVideoForm(
            request.POST,
            request.FILES,
        )

        if form.is_valid():

            try:

                output_path = process_resize(
                    video=form.cleaned_data["video"],
                    width=form.cleaned_data["width"],
                    height=form.cleaned_data["height"],
                    aspect_ratio=form.cleaned_data.get(
                        "aspect_ratio",
                        "",
                    ),
                    fit_mode=form.cleaned_data.get(
                        "fit_mode",
                        "fit",
                    ),
                    zoom=form.cleaned_data.get(
                        "zoom",
                        1.0,
                    ),
                    position_x=form.cleaned_data.get(
                        "position_x",
                        0,
                    ),
                    position_y=form.cleaned_data.get(
                        "position_y",
                        0,
                    ),
                    background_color=form.cleaned_data.get(
                        "background_color",
                        "#000000",
                    ),
                    output_format=form.cleaned_data.get(
                        "output_format",
                        "mp4",
                    ),
                )

                output_url = (
                    settings.MEDIA_URL
                    + "video_tools/outputs/"
                    + output_path.name
                )

                messages.success(
                    request,
                    "Video resized successfully.",
                )

                return render(
                    request,
                    "media_tools/resize.html",
                    {
                        "form": ResizeVideoForm(),
                        "output_url": output_url,
                    },
                )

            except ValueError as exc:

                logger.warning(
                    "Resize validation error: %s",
                    exc,
                )

                form.add_error(
                    None,
                    str(exc),
                )

                messages.error(
                    request,
                    str(exc),
                )

            except Exception:

                logger.exception(
                    "Unexpected resize view error."
                )

                form.add_error(
                    None,
                    "Unable to process the video.",
                )

                messages.error(
                    request,
                    "Unable to process the video.",
                )

    else:

        form = ResizeVideoForm()

    return render(
        request,
        "media_tools/resize.html",
        {
            "form": form,
        },
    )