import json
from django.test import TestCase, Client
from django.contrib.auth import get_user_model
from django.urls import reverse
from services.models import Plan, Feature, PlanFeature, Subscription, UsageRecord, FEATURE_CODES

User = get_user_model()

class PlanFeatureTests(TestCase):
    def setUp(self):
        # Create test users
        self.user_monthly = User.objects.create_user(username='monthly', password='pw')
        self.user_yearly = User.objects.create_user(username='yearly', password='pw')
        self.superuser = User.objects.create_superuser(username='admin', password='pw')

        # Create plans
        self.plan_free = Plan.objects.create(code='free', name='FREE', monthly_price=5, yearly_price=50, is_active=True)
        self.plan_pro = Plan.objects.create(code='pro', name='PRO', monthly_price=100, yearly_price=1000, is_active=True)
        self.plan_biz = Plan.objects.create(code='business_plus', name='BUSINESS+', pricing_type='contact', is_active=True)

        # Create features
        self.feat_qr = Feature.objects.create(key='qr_code', name='QR Code', is_active=True)
        self.feat_analytics = Feature.objects.create(key='analytics', name='Analytics', is_active=True)

        # Set up PlanFeatures
        self.pf_free_qr = PlanFeature.objects.create(plan=self.plan_free, feature=self.feat_qr, enabled=True, monthly_limit=5, yearly_limit=50)
        self.pf_pro_qr = PlanFeature.objects.create(plan=self.plan_pro, feature=self.feat_qr, enabled=True, monthly_limit=100, yearly_limit=1000)
        self.pf_biz_qr = PlanFeature.objects.create(plan=self.plan_biz, feature=self.feat_qr, enabled=True, is_unlimited=True)

        # Subscriptions
        self.sub_monthly = Subscription.objects.create(user=self.user_monthly, plan=self.plan_free, status='Active', billing_cycle='monthly')
        self.sub_yearly = Subscription.objects.create(user=self.user_yearly, plan=self.plan_free, status='Active', billing_cycle='yearly')

        self.client = Client()

    def test_entitlement_service_selects_correct_limit(self):
        from services.plan_features import get_feature_status, can_use_feature, increment_feature_usage

        # Monthly user should get monthly limit (5)
        status_m = get_feature_status(self.user_monthly, 'qr_code')
        self.assertEqual(status_m['limit'], 5)
        self.assertEqual(status_m['billing_cycle'], 'monthly')

        # Yearly user should get yearly limit (50)
        status_y = get_feature_status(self.user_yearly, 'qr_code')
        self.assertEqual(status_y['limit'], 50)
        self.assertEqual(status_y['billing_cycle'], 'yearly')

        # Increment usage for monthly user
        increment_feature_usage(self.user_monthly, 'qr_code')
        status_m = get_feature_status(self.user_monthly, 'qr_code')
        self.assertEqual(status_m['used'], 1)
        self.assertEqual(status_m['remaining'], 4)

    def test_admin_save_features_endpoint(self):
        self.client.login(username='admin', password='pw')
        url = reverse('custom_admin:plan_save_features', args=[self.plan_free.id])

        payload = {
            'pricing_type': 'fixed',
            'monthly_price': 6,
            'yearly_price': 60,
            'features': {
                'qr_code': {
                    'enabled': True,
                    'is_unlimited': False,
                    'monthly_limit': 10,
                    'yearly_limit': 100,
                },
                'analytics': {
                    'enabled': True,
                    'is_unlimited': False,
                    'history_days': 30
                }
            }
        }

        response = self.client.post(url, data=json.dumps(payload), content_type='application/json')
        self.assertEqual(response.status_code, 200)
        
        # Verify db changes
        self.plan_free.refresh_from_db()
        self.assertEqual(self.plan_free.monthly_price, 6)

        self.pf_free_qr.refresh_from_db()
        self.assertEqual(self.pf_free_qr.monthly_limit, 10)
        self.assertEqual(self.pf_free_qr.yearly_limit, 100)

        pf_analytics = PlanFeature.objects.get(plan=self.plan_free, feature__key='analytics')
        self.assertEqual(pf_analytics.history_days, 30)

    def test_pricing_view_saving_percentage(self):
        # Free: 5/mo (60/yr), yearly is 50 -> saving = 16.6% ~ 17%
        # Pro: 100/mo (1200/yr), yearly is 1000 -> saving = 16.6% ~ 17%
        response = self.client.get(reverse('services:pricing'))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context['max_saving_percent'], 17)
        
    def test_pricing_syncs_with_db(self):
        # Admin updates PRO plan price
        self.plan_pro.monthly_price = 150
        self.plan_pro.yearly_price = 1500
        self.plan_pro.save()
        
        response = self.client.get(reverse('services:pricing'))
        self.assertEqual(response.status_code, 200)
        
        # Ensure public pricing views the exact DB value without caching or resetting
        plans = list(response.context['plans'])
        pro_plan = next(p for p in plans if p.code == 'pro')
        self.assertEqual(pro_plan.monthly_price, 150)
        self.assertEqual(pro_plan.yearly_price, 1500)
