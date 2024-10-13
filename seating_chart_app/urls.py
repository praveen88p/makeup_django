from django.urls import path
from seating_chart_app import views

urlpatterns = [
    path('video_feed', views.video_feed, name='video_feed'),
    path('', views.index, name='index'),
]
