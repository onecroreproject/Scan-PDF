from django.urls import path
from . import views

app_name = 'converter'

urlpatterns = [
    path('', views.home, name='home'),
    path('convert/<str:tool_slug>/', views.convert_page, name='convert_page'),
    path('api/convert/<str:tool_slug>/', views.convert_file, name='convert_file'),
    path('api/speedtest/download/', views.speedtest_download, name='speedtest_download'),
    path('api/speedtest/upload/', views.speedtest_upload, name='speedtest_upload'),
    path('api/speedtest/client-info/', views.get_client_info, name='get_client_info'),
    path('preview-404/', views.custom_404_view, name='preview_404'),
]
