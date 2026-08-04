
from django.shortcuts import render
from django.contrib import messages
from django.conf import settings

from .forms import VideoAudioForm

from .services.validation import FileValidator
from .services.video_processing import VideoProcessor
from .services.audio_processing import AudioProcessor
from .services.volume_control import VolumeController
from .services.audio_loop import AudioLooper
from .services.duration_control import DurationController
from .services.replace_audio import AudioReplacer
from .services.merge_audio import AudioMerger
from .services.output_video import OutputVideo


from django.conf import settings
from django.contrib import messages
from django.shortcuts import render
import traceback


def upload(request):

    # -----------------------------------
    # GET Request
    # -----------------------------------
    if request.method == "GET":

        form = VideoAudioForm()

        return render(
            request,
            "audio_replacement/upload.html",
            {
                "form": form
            }
        )

    # -----------------------------------
    # POST Request
    # -----------------------------------
    form = VideoAudioForm(
        request.POST,
        request.FILES
    )

    # -----------------------------------
    # Form Validation
    # -----------------------------------
    if not form.is_valid():

        messages.error(
            request,
            "Please upload both a video and an audio file."
        )

        return render(
            request,
            "audio_replacement/upload.html",
            {
                "form": form
            }
        )

    obj = form.save()

    try:

        print("Step 1 : Validate Files")

        FileValidator.validate(
            obj.video,
            obj.audio
        )

        print("Step 2 : Read Video")

        video_processor = VideoProcessor(
            obj.video.path
        )

        video_info = video_processor.get_video_info()

        video_duration = video_info["duration"]

        print("Step 3 : Process Audio")

        audio_processor = AudioProcessor(
            obj.audio.path
        )

        processed_audio = audio_processor.process()

        print("Step 4 : Volume")

        volume = VolumeController(
            processed_audio
        )

        processed_audio = volume.set_volume(
            obj.volume
        )

        print("Step 5 : Loop")

        if obj.loop_audio:

            looper = AudioLooper(
                processed_audio
            )

            processed_audio = looper.auto_loop(
                video_duration
            )

        print("Step 6 : Trim")

        if obj.end_time != "00:00:00":

            duration = DurationController(
                processed_audio
            )

            processed_audio = duration.trim_audio(
                obj.start_time,
                obj.end_time
            )

        print("Step 7 : Replace / Merge")

        if obj.mode == "replace":

            replacer = AudioReplacer(
                obj.video.path,
                processed_audio
            )

            generated_video = replacer.replace()

        else:

            merger = AudioMerger(
                obj.video.path,
                processed_audio
            )

            generated_video = merger.merge()

        print("Generated Video :", generated_video)

        print("Step 8 : Final Output")

        output = OutputVideo(
            generated_video
        )

        output_video = output.finalize()

        print("Output Video :", output_video)

        obj.output_video = output_video
        obj.status = "Completed"
        obj.save()

        messages.success(
            request,
            "Video processed successfully."
        )

        print("Rendering Result Page")

        return render(
            request,
            "audio_replacement/result.html",
            {
                "output_video": output_video,
                "video_info": video_info,
                "MEDIA_URL": settings.MEDIA_URL,
            }
        )

    except Exception as e:

        traceback.print_exc()

        print("=" * 80)
        print("PROCESS FAILED")
        print(e)
        print("=" * 80)

        obj.status = "Failed"
        obj.save()

        messages.error(
            request,
            str(e)
        )

        return render(
            request,
            "audio_replacement/upload.html",
            {
                "form": form
            }
        )


#====================Add text to video========================

import os

from django.conf import settings
from django.http import FileResponse
from django.shortcuts import render

from .forms import AddTextVideoForm
from .services.text.text_service import process_text_video


from django.conf import settings

from django.conf import settings
from django.contrib import messages



def add_text_to_video(request):

    if request.method == "POST":

        form = AddTextVideoForm(
            request.POST,
            request.FILES
        )

        # ----------------------------
        # Form Validation
        # ----------------------------
        if not form.is_valid():

            messages.error(
                request,
                "Please upload a video and enter the required text."
            )

            return render(
                request,
                "video_tools/add_text.html",
                {
                    "form": form
                }
            )

        try:

            output_video = process_text_video(
                form.cleaned_data
            )

            output_video_url = (
                settings.MEDIA_URL +
                output_video.replace(
                    settings.MEDIA_ROOT,
                    ""
                ).replace("\\", "/")
            )

            messages.success(
                request,
                "Text added to the video successfully."
            )

            return render(
                request,
                "video_tools/result.html",
                {
                    "video_url": output_video_url,
                }
            )

        except Exception as e:

            messages.error(
                request,
                str(e)
            )

            return render(
                request,
                "video_tools/add_text.html",
                {
                    "form": form
                }
            )

    form = AddTextVideoForm()

    return render(
        request,
        "video_tools/add_text.html",
        {
            "form": form
        }
    )