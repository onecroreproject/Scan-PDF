from django.urls import path
from . import views

app_name = 'audio_processor'

urlpatterns = [
    path('editor/', views.editor_page, name='editor'),
    path('api/process/', views.process_audio_api, name='process_api'),
]
