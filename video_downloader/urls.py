from django.urls import path
from . import views

app_name = 'video_downloader'

urlpatterns = [
    path('', views.index, name='index'),
    path('youtube/', views.youtube_downloader, name='youtube_downloader'),
    path('facebook/', views.facebook_downloader, name='facebook_downloader'),
    path('twitter/', views.twitter_downloader, name='twitter_downloader'),
    path('instagram/', views.instagram_downloader, name='instagram_downloader'),
    path('tiktok/', views.tiktok_downloader, name='tiktok_downloader'),
    path('vimeo/', views.vimeo_downloader, name='vimeo_downloader'),
    path('reddit/', views.reddit_downloader, name='reddit_downloader'),
    path('dailymotion/', views.dailymotion_downloader, name='dailymotion_downloader'),
    path('api/analyze/', views.analyze_url, name='analyze_url'),
    path('api/download/', views.download_video, name='download_video'),
]
