from django.urls import path
from . import views

app_name = 'image_processor'

urlpatterns = [
    path('tool/<slug:tool_slug>/', views.tool_page, name='tool_page'),
    path('process/<slug:tool_slug>/', views.process_tool, name='process_tool'),
]
