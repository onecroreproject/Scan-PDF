from django.db import migrations, transaction

def seed_base_shorturl_feature(apps, schema_editor):
    Plan = apps.get_model('services', 'Plan')
    Feature = apps.get_model('services', 'Feature')
    PlanFeature = apps.get_model('services', 'PlanFeature')

    with transaction.atomic():
        feature, created = Feature.objects.get_or_create(
            key='short_url',
            defaults={
                'name': 'Short URLs',
                'type': 'NUMERIC',
                'section': 'SHORT URL',
                'display_order': 0,
                'is_public': True,
                'is_active': True,
            }
        )

        for plan in Plan.objects.all():
            # If the plan has unlimited dynamic qrs/short urls semantically 
            # (usually mapped via pricing_type='contact' in this project)
            is_unlimited = plan.pricing_type == 'contact'

            # Default to max_short_urls if available and not 0, otherwise some reasonable default
            limit = plan.max_short_urls if getattr(plan, 'max_short_urls', 0) > 0 else 10
            
            # Create or update safely without overwriting if already manually tweaked
            # If it already exists, get_or_create won't touch it.
            PlanFeature.objects.get_or_create(
                plan=plan,
                feature=feature,
                defaults={
                    'enabled': True,
                    'is_unlimited': is_unlimited,
                    'monthly_limit': limit if not is_unlimited else None,
                    'yearly_limit': limit * 10 if not is_unlimited else None, # Legacy systems often used 10x for yearly
                }
            )

def reverse_seed_base_shorturl_feature(apps, schema_editor):
    Feature = apps.get_model('services', 'Feature')
    Feature.objects.filter(key='short_url').delete()

class Migration(migrations.Migration):

    dependencies = [
        ('services', '0002_seed_shorturl_features'),
    ]

    operations = [
        migrations.RunPython(seed_base_shorturl_feature, reverse_code=reverse_seed_base_shorturl_feature),
    ]
