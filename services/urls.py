from django.urls import path
from . import views

app_name = 'services'

urlpatterns = [
    path('pricing/', views.pricing_view, name='pricing'),
    path('payment/confirm/<str:plan_code>/<str:cycle>/', views.payment_confirm_view, name='payment_confirm'),
    path('payment/simulate/', views.payment_simulate_view, name='payment_simulate'),
    path('payment/success/', views.payment_success_view, name='payment_success_view'),
    path('payment/failed/', views.payment_failed_view, name='payment_failed_view'),
    path('payment/history/', views.payment_history_view, name='payment_history'),
    path('help/', views.help_view, name='help'),
    path('support/', views.support_legacy_redirect, name='support'),
    path('contact/', views.contact_view, name='contact'),
]
