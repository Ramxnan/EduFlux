from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
import os
import shutil
from django.conf import settings
import django

class Command(BaseCommand):
    help = 'Deletes all corresponding folders in media/storage for non-admin users'

    def handle(self, *args, **kwargs):
        if not settings.configured:
            os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'eduflux.settings')
            django.setup()
        
        for user in User.objects.filter(is_superuser=False):
            # Construct the path to the user's folder
            display_name = user.username.split('@')[0]
            user_folder = os.path.join(settings.MEDIA_ROOT, 'storage', display_name)
            
            # Check if the folder exists
            if os.path.exists(user_folder):
                # If it exists, delete the folder and all its contents
                shutil.rmtree(user_folder)
                self.stdout.write(self.style.SUCCESS(f'Deleted folder for user: {display_name}'))
        
        self.stdout.write(self.style.SUCCESS('Successfully deleted all folders for non-admin users'))
