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
    path('upload_multiple_files_branch/', views.upload_multiple_files_branch, name='upload_multiple_files_branch'),
    path('download_file/<str:file_name>/', views.download_file, name='download_file'),
    path('download_folder/<str:folder_name>/', views.download_folder, name='download_folder'),
    path('delete/<str:file_name>/', views.delete_file, name='delete_file'),
    path('delete_folder/<str:folder_name>/', views.delete_folder, name='delete_folder'),
    ]+ static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)