from django.core.management.base import BaseCommand
from services.models import Subscription
from services.signals import _create_subscription_snapshot

class Command(BaseCommand):
    help = 'Backfills SubscriptionSnapshots for existing active subscriptions'

    def handle(self, *args, **kwargs):
        active_subs = Subscription.objects.filter(status='Active', snapshot__isnull=True)
        count = 0
        for sub in active_subs:
            try:
                _create_subscription_snapshot(sub)
                count += 1
                self.stdout.write(self.style.SUCCESS(f"Created snapshot for {sub.id} - {sub.user.username}"))
            except Exception as e:
                self.stdout.write(self.style.ERROR(f"Error creating snapshot for {sub.id}: {str(e)}"))
                
        self.stdout.write(self.style.SUCCESS(f"Successfully backfilled {count} snapshots."))
