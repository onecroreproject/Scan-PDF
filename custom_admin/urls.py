from django.urls import path
from . import views

app_name = 'custom_admin'

urlpatterns = [
    path('', views.dashboard_view, name='dashboard'),
    path('users/', views.users_view, name='users'),
    path('users/<int:user_id>/', views.user_detail_view, name='user_detail'),
    path('subscriptions/', views.subscriptions_view, name='subscriptions'),
    path('payments/', views.payments_view, name='payments'),
    path('plans/', views.plans_view, name='plans'),
    path('plans/<int:plan_id>/edit/', views.plan_edit_view, name='plan_edit'),
    path('features/', views.features_view, name='features'),
    path('qrcodes/', views.qrcodes_view, name='qrcodes'),
    path('shorturls/', views.shorturls_view, name='shorturls'),
    path('reports/', views.reports_view, name='reports'),
    path('activity/', views.activity_view, name='activity'),
    path('settings/', views.settings_view, name='settings'),
    path('logout/', views.logout_view, name='logout'),
]
