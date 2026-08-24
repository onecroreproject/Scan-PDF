from django.urls import path

from . import views

app_name="media_tools"

urlpatterns =[  

    path(
    "trim-video/",
    views.trim_video_view,
    name="trim_video",
),

    path(
        "merge_video/",
        views.merge_video_view,
        name="merge_video",
    ),
    path(
        "crop-video/",
        views.crop_video_view,
        name="crop_video",
    ),


]