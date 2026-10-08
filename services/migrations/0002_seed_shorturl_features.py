from django.db import migrations, transaction

def seed_shorturl_features(apps, schema_editor):
    Plan = apps.get_model('services', 'Plan')
    Feature = apps.get_model('services', 'Feature')
    PlanFeature = apps.get_model('services', 'PlanFeature')

    FEATURES_DATA = [
        {'key': 'header', 'name': 'Header', 'type': 'NUMERIC', 'free': {'m': 5, 'y': 50}, 'pro': {'m': 100, 'y': 1000}},
        {'key': 'qr_code', 'name': 'QR Code', 'type': 'NUMERIC', 'free': {'m': 5, 'y': 50}, 'pro': {'m': 100, 'y': 1000}},
        {'key': 'password_protection', 'name': 'Password Protection', 'type': 'NUMERIC', 'free': {'m': 3, 'y': 30}, 'pro': {'m': 50, 'y': 500}},
        {'key': 'link_expiry', 'name': 'Link Expiry', 'type': 'NUMERIC', 'free': {'m': 3, 'y': 30}, 'pro': {'m': 50, 'y': 500}},
        {'key': 'gps_tracking', 'name': 'GPS Tracking', 'type': 'NUMERIC', 'free': {'m': 2, 'y': 20}, 'pro': {'m': 50, 'y': 500}},
        {'key': 'analytics', 'name': 'Analytics', 'type': 'DURATION', 'free': {'d': 7}, 'pro': {'d': 365}},
        {'key': 'custom_alias', 'name': 'Custom Alias', 'type': 'NUMERIC', 'free': {'m': 2, 'y': 20}, 'pro': {'m': 50, 'y': 500}},
        {'key': 'csv_export', 'name': 'CSV Export', 'type': 'NUMERIC', 'free': {'m': 2, 'y': 20}, 'pro': {'m': 50, 'y': 500}},
        {'key': 'pdf_report', 'name': 'PDF Report', 'type': 'NUMERIC', 'free': {'m': 2, 'y': 20}, 'pro': {'m': 50, 'y': 500}},
        {'key': 'shorturl_utm', 'name': 'UTM Parameters', 'type': 'NUMERIC', 'free': {'m': 5, 'y': 50}, 'pro': {'m': 50, 'y': 500}},
        {'key': 'shorturl_cloaking', 'name': 'URL Cloaking', 'type': 'NUMERIC', 'free': {'m': 5, 'y': 50}, 'pro': {'m': 50, 'y': 500}},
    ]

    with transaction.atomic():
        # Find Plans safely
        free_plan = Plan.objects.filter(code='free').first() or Plan.objects.filter(name__iexact='free').first()
        pro_plan = Plan.objects.filter(code='pro').first() or Plan.objects.filter(name__iexact='pro').first()
        business_plan = Plan.objects.filter(code='business_plus').first() or Plan.objects.filter(name__iexact='business+').first()

        if not free_plan:
            free_plan = Plan.objects.create(name='FREE', code='free', monthly_price=0, yearly_price=0)

        if not pro_plan:
            pro_plan = Plan.objects.create(name='PRO', code='pro', monthly_price=100, yearly_price=1000)

        if not business_plan:
            business_plan = Plan.objects.create(name='BUSINESS+', code='business_plus', pricing_type='contact')

        for i, data in enumerate(FEATURES_DATA, start=1):
            feature_key = data['key']
            feature_name = data['name']
            feature_type = data['type']

            feature, created = Feature.objects.get_or_create(
                key=feature_key,
                defaults={
                    'name': feature_name,
                    'type': feature_type,
                    'section': 'SHORT URL',
                    'display_order': i,
                    'is_public': True,
                    'is_active': True,
                }
            )

            # Free Plan Limits
            defaults_free = {'enabled': True, 'is_unlimited': False}
            if 'd' in data['free']:
                defaults_free['history_days'] = data['free']['d']
            else:
                defaults_free['monthly_limit'] = data['free']['m']
                defaults_free['yearly_limit'] = data['free']['y']

            PlanFeature.objects.get_or_create(
                plan=free_plan,
                feature=feature,
                defaults=defaults_free
            )

            # Pro Plan Limits
            defaults_pro = {'enabled': True, 'is_unlimited': False}
            if 'd' in data['pro']:
                defaults_pro['history_days'] = data['pro']['d']
            else:
                defaults_pro['monthly_limit'] = data['pro']['m']
                defaults_pro['yearly_limit'] = data['pro']['y']

            PlanFeature.objects.get_or_create(
                plan=pro_plan,
                feature=feature,
                defaults=defaults_pro
            )

            # Business+ Limits
            defaults_bus = {'enabled': True, 'is_unlimited': True}
            PlanFeature.objects.get_or_create(
                plan=business_plan,
                feature=feature,
                defaults=defaults_bus
            )

def reverse_seed_shorturl_features(apps, schema_editor):
    pass

class Migration(migrations.Migration):

    dependencies = [
        ('services', '0001_initial'),
    ]

    operations = [
        migrations.RunPython(seed_shorturl_features, reverse_code=reverse_seed_shorturl_features),
    ]
