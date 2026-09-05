# Generated manually to fix legacy limits and prices

from django.db import migrations

def fix_plan_data(apps, schema_editor):
    Plan = apps.get_model('services', 'Plan')
    PlanFeature = apps.get_model('services', 'PlanFeature')

    # 1. Update Plans
    Plan.objects.filter(code='free').update(
        monthly_price=5, yearly_price=50, pricing_type='fixed'
    )
    Plan.objects.filter(code='pro').update(
        monthly_price=100, yearly_price=1000, pricing_type='fixed'
    )
    Plan.objects.filter(code='business_plus').update(
        monthly_price=0, yearly_price=0, pricing_type='contact'
    )

    # 2. Update PlanFeatures
    # FREE Plan
    free_plan = Plan.objects.filter(code='free').first()
    if free_plan:
        config = {
            'header': (5, 50),
            'qr_code': (5, 50),
            'password_protection': (3, 30),
            'link_expiry': (3, 30),
            'gps_tracking': (2, 20),
            'custom_alias': (2, 20),
            'csv_export': (2, 20),
            'pdf_report': (2, 20),
        }
        for code, (m, y) in config.items():
            PlanFeature.objects.filter(plan=free_plan, feature__key=code).update(
                monthly_limit=m, yearly_limit=y, is_unlimited=False
            )
        # Analytics
        PlanFeature.objects.filter(plan=free_plan, feature__key='analytics').update(
            history_days=7, is_unlimited=False
        )

    # PRO Plan
    pro_plan = Plan.objects.filter(code='pro').first()
    if pro_plan:
        config_pro = {
            'header': (100, 1000),
            'qr_code': (100, 1000),
            'password_protection': (50, 500),
            'link_expiry': (50, 500),
            'gps_tracking': (50, 500),
            'custom_alias': (50, 500),
            'csv_export': (50, 500),
            'pdf_report': (50, 500),
        }
        for code, (m, y) in config_pro.items():
            PlanFeature.objects.filter(plan=pro_plan, feature__key=code).update(
                monthly_limit=m, yearly_limit=y, is_unlimited=False
            )
        # Analytics
        PlanFeature.objects.filter(plan=pro_plan, feature__key='analytics').update(
            history_days=365, is_unlimited=False
        )

    # BUSINESS+ Plan
    biz_plan = Plan.objects.filter(code='business_plus').first()
    if biz_plan:
        PlanFeature.objects.filter(plan=biz_plan).update(
            is_unlimited=True, monthly_limit=None, yearly_limit=None, history_days=None
        )


class Migration(migrations.Migration):

    dependencies = [
        ('services', '0013_remove_planfeature_limit_and_more'),
    ]

    operations = [
        migrations.RunPython(fix_plan_data, reverse_code=migrations.RunPython.noop),
    ]
