from django.urls import path
from . import views

app_name = 'video_processor'

urlpatterns = [
    path('converter/', views.converter_page, name='converter'),
    path('api/convert/', views.convert_video_api, name='convert_api'),
]
