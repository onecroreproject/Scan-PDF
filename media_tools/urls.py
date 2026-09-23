from django.urls import path

from . import video_views

app_name = "media_tools"

#need every services file has when user select upload butoon when js come  in front

urlpatterns = [
    # Existing 21 Video Modules (with aliases)
    path("trim/", video_views.trim_video, name="trim"),#work

    path("merge/", video_views.merge_video, name="merge"),

    path("crop/", video_views.crop_video, name="crop"),

    path("resize/", video_views.resize_video, name="resize"),#option remove manual dimension and strech

    path("rotate/", video_views.rotate_video, name="rotate"),#invalid rotate angle come
    path("flip/", video_views.flip_video, name="flip"),
    path("speed/", video_views.speed_video, name="speed"),#work
    path("compress/", video_views.compress_video, name="compress"),#work
    path("convert/", video_views.convert_video, name="convert"),#work priview didnt show
    path("video-to-gif/", video_views.video_to_gif, name="video_to_gif"),#check
    path("gif-to-video/", video_views.gif_to_video, name="gif_to_video"),#check
    path("remove-audio/", video_views.remove_audio, name="remove_audio"),
    path("extract-audio/", video_views.extract_audio, name="extract_audio"),#when user extract button click when download audio come want
    path("change-volume/", video_views.change_volume, name="change_volume"),#work
    path("reverse/", video_views.reverse_video, name="reverse"),
    path("fade/", video_views.fade_video, name="fade"),
    path("effects/", video_views.effects_video, name="effects"),
    path("fps/", video_views.fps_video, name="fps"),
    path("loop/", video_views.loop_video, name="loop"),
    path("extract-frame/", video_views.extract_frame, name="extract_frame"),
    path("watermark/", video_views.watermark_video, name="watermark"),


    # New 11 Video Modules
    path("add-audio/", video_views.add_audio, name="add_audio"),
    path("add-text/", video_views.add_text, name="add_text"),
    path("add-image/", video_views.add_image, name="add_image"),
    path("blur/", video_views.blur_video, name="blur"),
    path("brighten/", video_views.brighten_video, name="brighten"),
    path("split-screen/", video_views.split_screen, name="split_screen"),
    path("stabilize/", video_views.stabilize_video, name="stabilize"),
    path("sync-audio/", video_views.sync_audio, name="sync_audio"),
    path("transition/", video_views.transition_video, name="transition"),
    path("repair/", video_views.repair_video, name="repair"),
    path("background/", video_views.background_video, name="background"),
    path("tool/<str:tool_name>/", video_views.video_tool_view, name="tool_generic"),
    path("download/<str:filename>/", video_views.download_output, name="download_output"),
]

