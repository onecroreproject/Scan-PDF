"""
Migration 0011 — Seed Feature master records, PlanFeature records, and update Plan prices.

This is the canonical data migration that establishes the new entitlement system.
All runtime feature checks will use the PlanFeature records created here.

Features seeded (9 Short URL features):
  header, qr_code, password_protection, link_expiry, gps_tracking,
  analytics, custom_alias, csv_export, pdf_report

Plans updated:
  FREE: monthly=5, yearly=50
  PRO: monthly=100, yearly=1000
  BUSINESS+: pricing_type='contact'
"""

from django.db import migrations


FEATURE_DEFINITIONS = [
    # (key, name, section, display_order)
    ('header',              'Header',               'SHORT URL', 1),
    ('qr_code',             'QR Code',              'SHORT URL', 2),
    ('password_protection', 'Password Protection',  'SHORT URL', 3),
    ('link_expiry',         'Link Expiry',          'SHORT URL', 4),
    ('gps_tracking',        'GPS Tracking',         'SHORT URL', 5),
    ('analytics',           'Analytics',            'SHORT URL', 6),
    ('custom_alias',        'Custom Alias',         'SHORT URL', 7),
    ('csv_export',          'CSV Export',           'SHORT URL', 8),
    ('pdf_report',          'PDF Report',           'SHORT URL', 9),
]

# PlanFeature configs per plan
# (feature_key, enabled, limit, is_unlimited, history_days)
PLAN_FEATURE_CONFIGS = {
    'free': [
        ('header',              True,  5,    False, None),
        ('qr_code',             True,  5,    False, None),
        ('password_protection', True,  3,    False, None),
        ('link_expiry',         True,  3,    False, None),
        ('gps_tracking',        True,  2,    False, None),
        ('analytics',           True,  None, False, 7),
        ('custom_alias',        True,  2,    False, None),
        ('csv_export',          True,  2,    False, None),
        ('pdf_report',          True,  2,    False, None),
    ],
    'pro': [
        ('header',              True,  100,  False, None),
        ('qr_code',             True,  100,  False, None),
        ('password_protection', True,  50,   False, None),
        ('link_expiry',         True,  50,   False, None),
        ('gps_tracking',        True,  50,   False, None),
        ('analytics',           True,  None, False, 365),
        ('custom_alias',        True,  50,   False, None),
        ('csv_export',          True,  50,   False, None),
        ('pdf_report',          True,  50,   False, None),
    ],
    'business_plus': [
        ('header',              True,  None, True,  None),
        ('qr_code',             True,  None, True,  None),
        ('password_protection', True,  None, True,  None),
        ('link_expiry',         True,  None, True,  None),
        ('gps_tracking',        True,  None, True,  None),
        ('analytics',           True,  None, True,  None),  # history_days=None means unlimited
        ('custom_alias',        True,  None, True,  None),
        ('csv_export',          True,  None, True,  None),
        ('pdf_report',          True,  None, True,  None),
    ],
}

PLAN_PRICE_UPDATES = {
    'free':         {'monthly_price': 5,   'yearly_price': 50,   'pricing_type': 'fixed', 'is_default': True},
    'pro':          {'monthly_price': 100, 'yearly_price': 1000, 'pricing_type': 'fixed'},
    'business_plus':{'monthly_price': 0,   'yearly_price': 0,    'pricing_type': 'contact'},
}


def seed_features_and_plan_features(apps, schema_editor):
    Plan = apps.get_model('services', 'Plan')
    Feature = apps.get_model('services', 'Feature')
    PlanFeature = apps.get_model('services', 'PlanFeature')

    # 1. Deactivate old BUSINESS plan
    Plan.objects.filter(code='business').update(is_active=False)

    # 2. Ensure the 3 required plans exist
    plan_defaults = {
        'free': {
            'name': 'FREE', 'monthly_price': 5, 'yearly_price': 50,
            'pricing_type': 'fixed', 'is_active': True, 'display_order': 1,
            'is_default': True, 'description': 'Perfect to get started',
        },
        'pro': {
            'name': 'PRO', 'monthly_price': 100, 'yearly_price': 1000,
            'pricing_type': 'fixed', 'is_active': True, 'display_order': 2,
            'is_popular': True, 'description': 'For serious creators',
        },
        'business_plus': {
            'name': 'BUSINESS+', 'monthly_price': 0, 'yearly_price': 0,
            'pricing_type': 'contact', 'is_active': True, 'display_order': 3,
            'description': 'Enterprise scale',
        },
    }
    for code, defaults in plan_defaults.items():
        Plan.objects.update_or_create(code=code, defaults=defaults)

    # 3. Create/update Feature master records
    feature_objs = {}
    for key, name, section, order in FEATURE_DEFINITIONS:
        feat, _ = Feature.objects.update_or_create(
            key=key,
            defaults={
                'name': name,
                'type': 'NUMERIC',
                'section': section,
                'is_public': True,
                'is_active': True,
                'display_order': order,
            }
        )
        feature_objs[key] = feat

    # 4. Create/update PlanFeature records
    for plan_code, feature_list in PLAN_FEATURE_CONFIGS.items():
        try:
            plan = Plan.objects.get(code=plan_code)
        except Plan.DoesNotExist:
            continue

        for feat_key, enabled, limit, is_unlimited, history_days in feature_list:
            feat = feature_objs.get(feat_key)
            if not feat:
                continue

            PlanFeature.objects.update_or_create(
                plan=plan,
                feature=feat,
                defaults={
                    'enabled': enabled,
                    'limit': limit,
                    'is_unlimited': is_unlimited,
                    'limit_period': 'monthly',
                    'history_days': history_days,
                    # Legacy compat fields
                    'value_boolean': enabled,
                    'value_numeric': limit,
                }
            )


def reverse_seed(apps, schema_editor):
    """Reverse: delete all the seeded PlanFeature and Feature records."""
    Feature = apps.get_model('services', 'Feature')
    PlanFeature = apps.get_model('services', 'PlanFeature')
    feature_keys = [f[0] for f in FEATURE_DEFINITIONS]
    PlanFeature.objects.filter(feature__key__in=feature_keys).delete()
    Feature.objects.filter(key__in=feature_keys).delete()


class Migration(migrations.Migration):

    dependencies = [
        ('services', '0010_add_plan_feature_entitlement_fields'),
    ]

    operations = [
        migrations.RunPython(seed_features_and_plan_features, reverse_code=reverse_seed),
    ]
