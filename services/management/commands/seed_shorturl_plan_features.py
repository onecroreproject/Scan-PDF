from django.core.management.base import BaseCommand
from services.models import Plan, Feature, PlanFeature
from django.db import transaction

class Command(BaseCommand):
    help = 'Seeds 11 Short URL features and PlanFeatures for Free, Pro, and Business+ plans.'

    def handle(self, *args, **options):
        # The exact 11 features defined by the user
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
                self.stdout.write(self.style.WARNING("FREE plan not found. Creating a placeholder."))
                free_plan = Plan.objects.create(name='FREE', code='free', monthly_price=0, yearly_price=0)

            if not pro_plan:
                self.stdout.write(self.style.WARNING("PRO plan not found. Creating a placeholder."))
                pro_plan = Plan.objects.create(name='PRO', code='pro', monthly_price=100, yearly_price=1000)

            if not business_plan:
                self.stdout.write(self.style.WARNING("BUSINESS+ plan not found. Creating a placeholder."))
                business_plan = Plan.objects.create(name='BUSINESS+', code='business_plus', pricing_type='contact')

            created_features_count = 0
            created_planfeatures_count = 0

            # Loop through 11 features
            for i, data in enumerate(FEATURES_DATA, start=1):
                feature_key = data['key']
                feature_name = data['name']
                feature_type = data['type']

                # 1. Get or create the Feature record
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
                if created:
                    created_features_count += 1
                else:
                    # Optional: update display_order or section just to ensure correct order
                    if feature.display_order != i or feature.section != 'SHORT URL' or feature.name != feature_name:
                        feature.display_order = i
                        feature.section = 'SHORT URL'
                        feature.name = feature_name
                        feature.save()

                # 2. Free Plan Limits (Safely get_or_create to NOT overwrite manual admin changes)
                defaults_free = {'enabled': True, 'is_unlimited': False}
                if 'd' in data['free']:
                    defaults_free['history_days'] = data['free']['d']
                else:
                    defaults_free['monthly_limit'] = data['free']['m']
                    defaults_free['yearly_limit'] = data['free']['y']

                pf_free, created_free = PlanFeature.objects.get_or_create(
                    plan=free_plan,
                    feature=feature,
                    defaults=defaults_free
                )
                if created_free:
                    created_planfeatures_count += 1

                # 3. Pro Plan Limits
                defaults_pro = {'enabled': True, 'is_unlimited': False}
                if 'd' in data['pro']:
                    defaults_pro['history_days'] = data['pro']['d']
                else:
                    defaults_pro['monthly_limit'] = data['pro']['m']
                    defaults_pro['yearly_limit'] = data['pro']['y']

                pf_pro, created_pro = PlanFeature.objects.get_or_create(
                    plan=pro_plan,
                    feature=feature,
                    defaults=defaults_pro
                )
                if created_pro:
                    created_planfeatures_count += 1

                # 4. Business+ Limits (Always Unlimited)
                defaults_bus = {'enabled': True, 'is_unlimited': True}
                pf_bus, created_bus = PlanFeature.objects.get_or_create(
                    plan=business_plan,
                    feature=feature,
                    defaults=defaults_bus
                )
                if created_bus:
                    created_planfeatures_count += 1

            self.stdout.write(self.style.SUCCESS(
                f"Successfully seeded! Created {created_features_count} new features and {created_planfeatures_count} new PlanFeature rows."
            ))
            
            # Verify database counts
            free_count = free_plan.plan_features.count()
            pro_count = pro_plan.plan_features.count()
            bus_count = business_plan.plan_features.count()
            self.stdout.write(self.style.SUCCESS(f"Current Feature Count -> FREE: {free_count}, PRO: {pro_count}, BUSINESS+: {bus_count}"))
