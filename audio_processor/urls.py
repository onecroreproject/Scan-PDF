from django.urls import path
from . import views

app_name = 'audio_processor'

urlpatterns = [
    path('editor/', views.editor_page, name='editor'),
    path('merge/', views.merge_page, name='merge'),
    path('extract/', views.extract_page, name='extract'),
    path('api/process/', views.process_audio_api, name='process_api'),
]
