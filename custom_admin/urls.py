from django.urls import path
from . import views

app_name = 'custom_admin'

urlpatterns = [
    # Dashboard
    path('', views.dashboard_view, name='dashboard'),

    # Users
    path('users/', views.users_view, name='users'),
    path('users/<int:user_id>/', views.user_detail_view, name='user_detail'),

    # Subscriptions & Payments
    path('subscriptions/', views.subscriptions_view, name='subscriptions'),
    path('payments/', views.payments_view, name='payments'),

    # Plans & Pricing
    path('plans/', views.plans_view, name='plans'),
    path('plans/<int:plan_id>/edit/', views.plan_edit_view, name='plan_edit'),

    # Section AJAX endpoints
    path('plans/<int:plan_id>/sections/create/', views.section_create, name='section_create'),
    path('sections/<int:section_id>/update/', views.section_update, name='section_update'),
    path('sections/<int:section_id>/delete/', views.section_delete, name='section_delete'),
    path('plans/<int:plan_id>/sections/reorder/', views.section_reorder, name='section_reorder'),

    # Feature AJAX endpoints
    path('sections/<int:section_id>/features/create/', views.feature_create, name='feature_create'),
    path('features/<int:feature_id>/update/', views.feature_update, name='feature_update'),
    path('features/<int:feature_id>/delete/', views.feature_delete, name='feature_delete'),
    path('sections/<int:section_id>/features/reorder/', views.feature_reorder, name='feature_reorder'),

    # Legacy features (redirects to plans)
    path('features/', views.features_view, name='features'),

    # Other
    path('qrcodes/', views.qrcodes_view, name='qrcodes'),
    path('shorturls/', views.shorturls_view, name='shorturls'),
    path('reports/', views.reports_view, name='reports'),
    path('activity/', views.activity_view, name='activity'),
    path('settings/', views.settings_view, name='settings'),
    path('logout/', views.logout_view, name='logout'),
]
