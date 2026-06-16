from django.urls import path
from . import views

app_name = 'video_processor'

urlpatterns = [
    # Tool pages
    path('converter/', views.converter_page, name='converter'),
    path('image-to-video/', views.image_to_video_page, name='image_to_video'),
    path('editor/', views.video_editor_page, name='video_editor'),
    path('compressor/', views.compressor_page, name='compressor'),
    path('merger/', views.merger_page, name='merger'),
    path('trimmer/', views.trimmer_page, name='trimmer'),
    path('gif-maker/', views.gif_maker_page, name='gif_maker'),
    path('audio-extractor/', views.audio_extractor_page, name='audio_extractor'),
    path('watermark/', views.watermark_page, name='watermark'),
    path('subtitle-overlay/', views.subtitle_overlay_page, name='subtitle_overlay'),

    # Chunk upload
    path('api/chunk-upload/', views.chunk_upload, name='chunk_upload'),
    path('api/chunk-status/', views.chunk_status, name='chunk_status'),

    # APIs
    path('api/convert/', views.convert_video_api, name='convert_api'),
    path('api/image-to-video/', views.image_to_video_api, name='image_to_video_api'),
    path('api/trim/', views.trim_video_api, name='trim_api'),
    path('api/cut/', views.cut_video_api, name='cut_api'),
    path('api/rotate/', views.rotate_video_api, name='rotate_api'),
    path('api/resize/', views.resize_video_api, name='resize_api'),
    path('api/crop/', views.crop_video_api, name='crop_api'),
    path('api/speed/', views.speed_video_api, name='speed_api'),
    path('api/mute/', views.mute_video_api, name='mute_api'),
    path('api/replace-audio/', views.replace_audio_api, name='replace_audio_api'),
    path('api/text-overlay/', views.text_overlay_api, name='text_overlay_api'),
    path('api/compress/', views.compress_video_api, name='compress_api'),
    path('api/compress-status/', views.compress_task_status, name='compress_status'),
    path('api/merge/', views.merge_videos_api, name='merge_api'),
    path('api/gif/', views.make_gif_api, name='gif_api'),
    path('api/extract-audio/', views.extract_audio_api, name='extract_audio_api'),
    path('api/watermark/', views.add_watermark_api, name='watermark_api'),
    path('api/subtitle/', views.add_subtitle_api, name='subtitle_api'),
    path('api/video-info/', views.video_info_api, name='video_info_api'),
]
