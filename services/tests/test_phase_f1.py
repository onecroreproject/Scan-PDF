import datetime
from django.utils import timezone
from django.test import TestCase
from django.contrib.auth import get_user_model
from django.db import transaction, IntegrityError
import threading

from services.models import Plan, Feature, PlanFeature, Subscription, UsageRecord, UsageOverride
from services.plan_features import (
    check_and_increment_short_url_features,
    _get_effective_limit,
    _get_billing_period,
    increment_feature_usage
)
from dynamic_qr.models import DynamicQRCode

User = get_user_model()

class PhaseF1QuotaTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_user(username='phase-f1-user', password='pw')
        self.plan = Plan.objects.create(code='f1-test', name='F1 Test Plan', is_active=True)
        
        # Features
        self.f_short_url, _ = Feature.objects.get_or_create(key='short_url', defaults={'name': 'Short URL'})
        self.f_qr, _ = Feature.objects.get_or_create(key='qr_code', defaults={'name': 'QR Code'})
        self.f_alias, _ = Feature.objects.get_or_create(key='custom_alias', defaults={'name': 'Custom Alias'})

        # PlanFeatures (limits)
        PlanFeature.objects.create(plan=self.plan, feature=self.f_short_url, enabled=True, monthly_limit=5)
        PlanFeature.objects.create(plan=self.plan, feature=self.f_qr, enabled=True, monthly_limit=2)
        PlanFeature.objects.create(plan=self.plan, feature=self.f_alias, enabled=True, monthly_limit=1)

        self.sub = Subscription.objects.create(
            user=self.user, plan=self.plan, status='Active', billing_cycle='monthly',
            start_date=timezone.now()
        )

    def test_base_quota_consumption(self):
        # 1. Plain Short URL (only short_url)
        ok, err_code, _ = check_and_increment_short_url_features(self.user, {'short_url': True})
        self.assertTrue(ok)
        self.assertEqual(UsageRecord.objects.get(user=self.user, feature_key='short_url').current_usage, 1)

        # 2. QR Short URL (short_url + qr_code)
        ok, err_code, _ = check_and_increment_short_url_features(self.user, {'short_url': True, 'qr_code': True})
        self.assertTrue(ok)
        self.assertEqual(UsageRecord.objects.get(user=self.user, feature_key='short_url').current_usage, 2)
        self.assertEqual(UsageRecord.objects.get(user=self.user, feature_key='qr_code').current_usage, 1)

        # 3. Alias + QR (short_url + qr_code + custom_alias)
        ok, err_code, _ = check_and_increment_short_url_features(self.user, {'short_url': True, 'qr_code': True, 'custom_alias': True})
        self.assertTrue(ok)
        self.assertEqual(UsageRecord.objects.get(user=self.user, feature_key='short_url').current_usage, 3)
        self.assertEqual(UsageRecord.objects.get(user=self.user, feature_key='qr_code').current_usage, 2)
        self.assertEqual(UsageRecord.objects.get(user=self.user, feature_key='custom_alias').current_usage, 1)

        # 4. Limit Exhausted
        ok, err_code, _ = check_and_increment_short_url_features(self.user, {'short_url': True, 'custom_alias': True})
        self.assertFalse(ok)
        self.assertEqual(err_code, 'feature_limit_reached')

        # 5. Edit an existing QR (should NOT consume short_url)
        existing = DynamicQRCode.objects.create(user=self.user, qr_type='url', destination_url='http://a.com')
        ok, err_code, _ = check_and_increment_short_url_features(self.user, {'short_url': True}, existing_qr=existing)
        self.assertTrue(ok)
        self.assertEqual(UsageRecord.objects.get(user=self.user, feature_key='short_url').current_usage, 3) # Still 3

    def test_atomicity_partial_exhaustion(self):
        # Base available (3/5), Alias exhausted (1/1)
        UsageRecord.objects.create(
            user=self.user, feature_key='custom_alias', period_start=self.sub.start_date,
            period_end=self.sub.start_date + datetime.timedelta(days=30), current_usage=1
        )
        ok, err_code, _ = check_and_increment_short_url_features(self.user, {'short_url': True, 'custom_alias': True})
        self.assertFalse(ok)
        # short_url should NOT have incremented
        self.assertFalse(UsageRecord.objects.filter(user=self.user, feature_key='short_url').exists())

    def test_override_logic(self):
        # Base limit = 5
        # 1. Additional Allowance
        UsageOverride.objects.create(user=self.user, feature_key='short_url', additional_allowance=3)
        self.assertEqual(_get_effective_limit(self.user, 'short_url', 5), 8)

        # 2. Absolute Override 0
        o2 = UsageOverride.objects.create(user=self.user, feature_key='short_url', override_limit=0)
        self.assertEqual(_get_effective_limit(self.user, 'short_url', 5), 0)

        # 3. Expired Override
        o2.expires_at = timezone.now() - datetime.timedelta(days=1)
        o2.save()
        # Fallbacks to the first valid one (additional_allowance=3) -> 8
        self.assertEqual(_get_effective_limit(self.user, 'short_url', 5), 8)

        # 4. Multiple Overrides (Deterministic: Max absolute wins)
        o2.expires_at = timezone.now() + datetime.timedelta(days=1)
        o2.override_limit = 10
        o2.save()
        UsageOverride.objects.create(user=self.user, feature_key='short_url', override_limit=15)
        self.assertEqual(_get_effective_limit(self.user, 'short_url', 5), 15)

    def test_billing_period_arithmetic(self):
        # Subscription start: Jan 31
        start_date = timezone.make_aware(datetime.datetime(2023, 1, 31, 12, 0))
        sub = Subscription(start_date=start_date, billing_cycle='monthly')
        
        # Current time: Feb 28
        now = timezone.make_aware(datetime.datetime(2023, 2, 28, 15, 0))
        import unittest.mock
        with unittest.mock.patch('django.utils.timezone.now', return_value=now):
            p_start, p_end = _get_billing_period(sub)
            self.assertEqual(p_start.month, 2)
            self.assertEqual(p_start.day, 28) # Feb has 28 days

        # Subscription start: Feb 29 (Leap year)
        start_date = timezone.make_aware(datetime.datetime(2024, 2, 29, 12, 0))
        sub = Subscription(start_date=start_date, billing_cycle='yearly')
        now = timezone.make_aware(datetime.datetime(2025, 3, 1, 15, 0))
        with unittest.mock.patch('django.utils.timezone.now', return_value=now):
            p_start, p_end = _get_billing_period(sub)
            self.assertEqual(p_start.year, 2025)
            self.assertEqual(p_start.month, 2)
            self.assertEqual(p_start.day, 28) # Not a leap year
