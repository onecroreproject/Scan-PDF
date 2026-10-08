import datetime
from django.test import TestCase
from django.core.exceptions import ValidationError
from django.contrib.admin.sites import AdminSite
from django.contrib.auth import get_user_model

from services.models import Feature, PlanFeature, UsageOverride, UsageRecord, Plan
from services.admin import FeatureAdmin, UsageRecordAdmin
from services.checks import check_core_features_exist

User = get_user_model()

class MockRequest:
    def __init__(self, user):
        self.user = user

class PhaseGAdminSafetyTests(TestCase):
    def setUp(self):
        self.admin_user = User.objects.create_superuser('admin_g', 'admin@g.com', 'pw')
        self.site = AdminSite()
        self.feature_admin = FeatureAdmin(Feature, self.site)
        self.record_admin = UsageRecordAdmin(UsageRecord, self.site)
        
        self.f_core = Feature.objects.create(key='short_url', name='Base')
        self.f_custom = Feature.objects.create(key='custom_thing', name='Custom')

        self.plan = Plan.objects.create(code='g-plan', name='G Plan')

    def test_core_feature_cannot_be_deleted(self):
        request = MockRequest(self.admin_user)
        self.assertFalse(self.feature_admin.has_delete_permission(request, self.f_core))
        self.assertTrue(self.feature_admin.has_delete_permission(request, self.f_custom))

    def test_core_feature_key_is_readonly(self):
        request = MockRequest(self.admin_user)
        readonly = self.feature_admin.get_readonly_fields(request, self.f_core)
        self.assertIn('key', readonly)

    def test_plan_feature_validation_rejects_negative(self):
        pf = PlanFeature(plan=self.plan, feature=self.f_core, monthly_limit=-5)
        with self.assertRaises(ValidationError):
            pf.clean()

    def test_usage_override_rejects_negative(self):
        override = UsageOverride(user=self.admin_user, feature_key='short_url', override_limit=-10)
        with self.assertRaises(ValidationError):
            override.clean()

    def test_usage_record_admin_is_readonly(self):
        request = MockRequest(self.admin_user)
        self.assertFalse(self.record_admin.has_add_permission(request))
        
        readonly_fields = self.record_admin.get_readonly_fields(request)
        self.assertIn('current_usage', readonly_fields)

    def test_system_check_detects_missing_features(self):
        # By default, since we only created 'short_url' and 'custom_thing', the others are missing
        errors = check_core_features_exist(None)
        self.assertTrue(len(errors) > 0)
        # Verify it found qr_code missing
        missing_keys = [e.msg for e in errors]
        self.assertTrue(any("qr_code" in msg for msg in missing_keys))
