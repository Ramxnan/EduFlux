from django.core.management.base import BaseCommand
from django.contrib.auth.models import User

class Command(BaseCommand):
    help = 'Lists all users in the database'

    def handle(self, *args, **kwargs):
        users = User.objects.all()
        if users.exists():
            self.stdout.write(self.style.SUCCESS('Listing all users:'))
            for user in users:
                self.stdout.write(f'Email: {user.email}, Is Superuser: {user.is_superuser}')
        else:
            self.stdout.write(self.style.WARNING('No users found in the database.'))
