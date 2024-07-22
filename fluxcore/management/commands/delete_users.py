from django.core.management.base import BaseCommand
from django.contrib.auth.models import User

class Command(BaseCommand):
    help = 'Deletes all users from the database'

    def handle(self, *args, **kwargs):
        users = User.objects.all()
        if users.exists():
            for user in users:
                username = user.username
                user.delete()
                self.stdout.write(self.style.SUCCESS(f'Deleted user: {username}'))
            self.stdout.write(self.style.SUCCESS('Successfully deleted all users.'))
        else:
            self.stdout.write(self.style.WARNING('No users found in the database.'))
