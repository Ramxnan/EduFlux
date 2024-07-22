from django.core.management.base import BaseCommand
from django.contrib.auth.models import User

class Command(BaseCommand):
    help = 'Deletes all superusers from the database'

    def handle(self, *args, **kwargs):
        superusers = User.objects.filter(is_superuser=True)
        if superusers.exists():
            for user in superusers:
                username = user.username
                user.delete()
                self.stdout.write(self.style.SUCCESS(f'Deleted superuser: {username}'))
            self.stdout.write(self.style.SUCCESS('Successfully deleted all superusers.'))
        else:
            self.stdout.write(self.style.WARNING('No superusers found in the database.'))
