from django.urls import path
from . import views
from .views import send_mails_view


urlpatterns = [
    path('home/', views.home, name='home'),
    path('blocks/', views.blocks_view, name='blocks'),
    path('rooms/', views.rooms_view, name='rooms'),
    path('exam-slot/', views.exam_slot_view, name='exam_slot'),
    path('upload/', views.upload_view, name='upload'),
    path('seating/', views.seating_view, name='seating'),
    path('download/', views.download_seating, name='download_seating'),
    path('send-mails/', send_mails_view, name='send_mails'),



]
