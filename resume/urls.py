from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('build/', views.resume_form, name='resume_form'),
    path('display/<int:pk>/', views.resume_display, name='resume_display'),
    path('feedback/<int:pk>/', views.resume_feedback, name='resume_feedback'),
    path('download/docx/<int:pk>/', views.download_docx, name='download_docx'),
    path('upload/', views.upload_scan, name='upload_scan'),
]