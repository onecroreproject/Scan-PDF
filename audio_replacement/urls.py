from django.urls import path
from . import views

app_name = 'video_tools'

urlpatterns=[

   path( "", views.upload, name="upload" ),


    path(
        "add-text/",
        views.add_text_to_video,
        name="add_text_to_video"
    ),

path(
    "add-watermark/",
    views.add_watermark,
    name="add_watermark",
),


]