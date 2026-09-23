from django.urls import path
from . import views

app_name = 'dynamic_qr'

urlpatterns = [
    # Auth (only for dynamic QR feature)
    path('login/', views.dqr_login_view, name='login'),
    path('register/', views.dqr_register_view, name='register'),
    path('logout/', views.dqr_logout_view, name='logout'),
    path('forgot-password/', views.dqr_forgot_password_view, name='forgot_password'),
    path('verify-otp/', views.dqr_verify_otp_view, name='verify_otp'),
    path('reset-password/', views.dqr_reset_password_view, name='reset_password'),

    # Dashboard
    path('dashboard/', views.dqr_dashboard_view, name='dashboard'),
    path('plan/', views.dqr_plan_view, name='plan'),
    path('all/', views.dqr_all_qrs_view, name='all_qrs'),
    path('create/', views.dqr_create_view, name='create'),
    path('short-url/', views.dqr_short_url_view, name='short_url'),
    path('short-url/get/<uuid:qr_id>/', views.dqr_get_short_url_details, name='get_short_url'),
    path('short-url/analytics/<uuid:qr_id>/', views.dqr_short_url_analytics_view, name='short_url_analytics'),
    path('edit/<uuid:qr_id>/', views.dqr_edit_view, name='edit'),
    path('delete/<uuid:qr_id>/', views.dqr_delete_view, name='delete'),
    path('details/<uuid:qr_id>/', views.dqr_details_view, name='details'),
    path('analytics/<uuid:qr_id>/', views.dqr_analytics_view, name='analytics'),
    path('toggle-status/<uuid:qr_id>/', views.dqr_toggle_status, name='toggle_status'),
    path('download/<uuid:qr_id>/', views.dqr_download_view, name='download'),

    # QR redirect (short URL)
    path('r/<str:short_code>/', views.dqr_redirect_view, name='redirect'),
    path('r/<str:header>/<str:short_code>/', views.dqr_redirect_with_header_view, name='redirect_header'),

    # API
    path('api/auth-status/', views.dqr_auth_status, name='auth_status'),
    path('api/generate-image/', views.dqr_generate_image, name='generate_image'),
    path('api/check-username/', views.dqr_check_username_view, name='check_username'),
    path('api/check-email/', views.dqr_check_email_view, name='check_email'),
]
