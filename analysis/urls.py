from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload, name='upload'),
    path('download/', views.download_excel, name='download_excel'),
    path('clear/', views.clear_results, name='clear_results'),
]
