from django.urls import path
from . import views
from . import api_views

app_name = 'converter'

urlpatterns = [
    path('', views.home, name='home'),
    path('convert/<str:tool_slug>/', views.convert_page, name='convert_page'),
    path('api/convert/<str:tool_slug>/', views.convert_file, name='convert_file'),
    path('api/currency-rates/', views.currency_rates, name='currency_rates'),
    path('api/speedtest/download/', views.speedtest_download, name='speedtest_download'),
    path('api/speedtest/upload/', views.speedtest_upload, name='speedtest_upload'),
    path('api/speedtest/client-info/', views.get_client_info, name='get_client_info'),
    path('preview-404/', views.custom_404_view, name='preview_404'),
    # ── QR Code Generator REST API ──────────────────────────────────────────
    # Versioned endpoint (primary)
    path('api/v1/qr/generate/', api_views.qr_generate_api, name='api_v1_qr_generate'),
    # Legacy alias — kept for backward compatibility, same view
    path('api/qr/generate/', api_views.qr_generate_api, name='qr_generate_api_legacy'),
]

