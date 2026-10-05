from django.urls import path
from . import views

app_name = 'video_downloader'

urlpatterns = [
    path('instagram/', views.instagram_downloader, name='instagram_downloader'),
    path('twitter/', views.twitter_downloader, name='twitter_downloader'),
    path('facebook/', views.facebook_downloader, name='facebook_downloader'),
    path('youtube/', views.youtube_downloader, name='youtube_downloader'),
    path('threads/', views.threads_downloader, name='threads_downloader'),
    path('api/analyze/', views.analyze_url, name='analyze_url'),
    path('api/download/', views.download_media, name='download_media'),
    path('api/progress/', views.download_progress, name='download_progress'),
]
