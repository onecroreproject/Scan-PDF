from django.urls import path
from . import views

app_name = 'converter'

urlpatterns = [
    path('', views.home, name='home'),
    path('convert/<str:tool_slug>/', views.convert_page, name='convert_page'),
    path('api/convert/<str:tool_slug>/', views.convert_file, name='convert_file'),
]
