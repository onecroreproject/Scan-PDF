"""
URL configuration for allinone project.
"""
from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static

from django.views.generic.base import RedirectView
from django.contrib.staticfiles.storage import staticfiles_storage

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('converter.urls')),
    path('image/', include('image_processor.urls')),
    path('audio/', include('audio_processor.urls')),
    path("videotools/",include("audio_replacement.urls")),
    path('qr/', include('dynamic_qr.urls')),
    path('video-downloader/', include('video_downloader.urls')),
    path('services/', include('services.urls')),
    path('custom-admin/', include('custom_admin.urls')),
    path('favicon.ico', RedirectView.as_view(url=staticfiles_storage.url('favicon.ico'))),
    path('media_tools/', include('media_tools.urls')),
]

handler404 = 'converter.views.custom_404_view'

if settings.DEBUG:
    from django.contrib.staticfiles.urls import staticfiles_urlpatterns
    urlpatterns += staticfiles_urlpatterns()
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
