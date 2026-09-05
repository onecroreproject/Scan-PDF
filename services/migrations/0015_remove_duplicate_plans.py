# Generated manually to remove duplicate plan rows

from django.db import migrations
from django.db.models import Count

def remove_duplicate_plans(apps, schema_editor):
    Plan = apps.get_model('services', 'Plan')
    
    # Identify plan codes that have more than one row
    duplicate_codes = Plan.objects.values('code').annotate(count=Count('id')).filter(count__gt=1)
    
    for entry in duplicate_codes:
        code = entry['code']
        # Get all plans with this code, ordered by id (oldest first)
        plans = list(Plan.objects.filter(code=code).order_by('id'))
        
        # We'll keep the first one
        canonical_plan = plans[0]
        
        # For the duplicates, we want to migrate any subscriptions pointing to them
        # to the canonical plan, and then delete the duplicate.
        Subscription = apps.get_model('services', 'Subscription')
        PlanFeature = apps.get_model('services', 'PlanFeature')
        
        for duplicate in plans[1:]:
            Subscription.objects.filter(plan=duplicate).update(plan=canonical_plan)
            # We don't need to migrate PlanFeature rows, the canonical plan should have its own.
            # We can safely delete the duplicate's PlanFeatures and the duplicate plan itself.
            PlanFeature.objects.filter(plan=duplicate).delete()
            duplicate.delete()

class Migration(migrations.Migration):

    dependencies = [
        ('services', '0014_fix_plan_feature_limits'),
    ]

    operations = [
        migrations.RunPython(remove_duplicate_plans, reverse_code=migrations.RunPython.noop),
    ]
