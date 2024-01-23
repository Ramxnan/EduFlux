from django.urls import path
from . import views
from django.conf.urls.static import static
from django.conf import settings
from django.shortcuts import redirect

def redirect_to_dashboard(request):
    return redirect('dashboard')

urlpatterns = [

    path('', views.homepage, name=''),
    path('homepage/', views.homepage, name='homepage'),
    path('register/', views.register, name='register'),
    path('login/', views.login, name='login'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('logout', views.logout, name='logout'),
    path('submit/', views.submit, name='submit'),
    path('download/<str:file_name>/', views.download_file, name='download_file'),
    path('delete/<str:file_name>/', views.delete_file, name='delete_file'),
    path('fetch_file_lists/', views.fetch_file_lists, name='fetch_file_lists'),
    ]+ static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)