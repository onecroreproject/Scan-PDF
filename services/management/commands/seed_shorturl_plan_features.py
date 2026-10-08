from django.core.management.base import BaseCommand
from services.models import Plan, Feature, PlanFeature
from django.db import transaction

class Command(BaseCommand):
    help = 'Seeds 11 Short URL features and PlanFeatures for Free, Pro, and Business+ plans.'

    def handle(self, *args, **options):
        # The exact 11 features defined by the user
        self.stdout.write(self.style.WARNING(
            "DEPRECATION WARNING: This command is no longer needed."
        ))
        self.stdout.write(self.style.WARNING(
            "Short URL plan features are now automatically seeded during 'python manage.py migrate'."
        ))
        self.stdout.write(self.style.WARNING(
            "If you need to repair missing features, please use Django Admin."
        ))
