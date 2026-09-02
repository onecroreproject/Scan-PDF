from django.urls import path

from . import views


from .video_views import (
    crop_video,resize_video,
)

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

   
    path(
        "crop/",
        crop_video,
        name="crop",
    ),

path(
    "video/resize/",
    resize_video,
    name="resize_video",
),

]