from django.utils import timezone
from .models import Subscription, PlanFeature, UsageRecord, UsageOverride, Plan
from django.db import models
import datetime

class PlanLimitService:
    @staticmethod
    def get_active_subscription(user):
        """Returns the active subscription for the user, prioritizing 'Active' status."""
        if not user.is_authenticated:
            return None
        sub = Subscription.objects.filter(user=user, status='Active').first()
        if not sub:
            # Check if there is a pending subscription, but usually we fallback to a default free plan
            default_plan = Plan.objects.filter(is_default=True).first()
            if default_plan:
                # Mock a subscription object for free users if one doesn't exist
                return Subscription(user=user, plan=default_plan, status='Active', billing_cycle='monthly', start_date=timezone.now())
        return sub

    @staticmethod
    def check_feature_access(user, feature_key):
        """Checks if the user's active plan allows a specific feature."""
        sub = PlanLimitService.get_active_subscription(user)
        if not sub:
            return False
            
        plan_feature = PlanFeature.objects.filter(plan=sub.plan, feature__key=feature_key).first()
        if not plan_feature:
            return False
            
        # If it's a boolean feature, check its boolean value
        if plan_feature.feature.type == 'BOOLEAN':
            return plan_feature.value_boolean
            
        # For non-boolean features, existence and not being empty typically means access
        if plan_feature.feature.type == 'NUMERIC':
            return plan_feature.value_numeric is not None and plan_feature.value_numeric > 0
        if plan_feature.feature.type == 'UNLIMITED':
            return True
            
        return True

    @staticmethod
    def get_billing_period(subscription):
        """Calculates the current billing period start and end dates."""
        now = timezone.now()
        start_date = subscription.start_date
        
        if subscription.billing_cycle == 'monthly':
            # Calculate months difference
            months_diff = (now.year - start_date.year) * 12 + now.month - start_date.month
            
            # If current day is before start day, subtract a month
            if now.day < start_date.day:
                months_diff -= 1
                
            # Current period start
            period_start = start_date + datetime.timedelta(days=30 * months_diff) # Approximation
            period_end = period_start + datetime.timedelta(days=30)
        else: # yearly
            years_diff = now.year - start_date.year
            if now.month < start_date.month or (now.month == start_date.month and now.day < start_date.day):
                years_diff -= 1
                
            period_start = start_date.replace(year=start_date.year + years_diff)
            period_end = period_start.replace(year=period_start.year + 1)
            
        return period_start, period_end

    @staticmethod
    def check_usage_limit(user, feature_key, increment=False):
        """
        Checks if the user has reached their limit for the given feature.
        If increment is True, increments the usage counter.
        Returns a tuple: (allowed: bool, current_usage: int, limit: int/str, error_msg: str)
        """
        sub = PlanLimitService.get_active_subscription(user)
        if not sub:
            return False, 0, 0, "No active subscription found."
            
        plan_feature = PlanFeature.objects.filter(plan=sub.plan, feature__key=feature_key).first()
        if not plan_feature:
            return False, 0, 0, "Feature not available in your plan."
            
        # If the feature is strictly boolean and not numeric
        if plan_feature.feature.type == 'BOOLEAN':
            if plan_feature.value_boolean:
                return True, 0, 'Unlimited', None
            else:
                return False, 0, 0, "Feature not enabled in your plan."
                
        # If unlimited
        if plan_feature.feature.type == 'UNLIMITED':
            # We might still want to track usage, but we never block
            if increment:
                PlanLimitService._increment_usage(user, feature_key, sub)
            # Fetch current usage for display
            current = PlanLimitService._get_current_usage(user, feature_key, sub)
            return True, current, 'Unlimited', None
            
        # If numeric limit
        if plan_feature.feature.type == 'NUMERIC':
            base_limit = plan_feature.value_numeric or 0
            
            # Check for overrides
            override = UsageOverride.objects.filter(user=user, feature_key=feature_key).filter(
                models.Q(expires_at__isnull=True) | models.Q(expires_at__gt=timezone.now())
            ).first()
            
            if override:
                if override.override_limit is not None:
                    base_limit = override.override_limit
                else:
                    base_limit += override.additional_allowance
                    
            current_usage = PlanLimitService._get_current_usage(user, feature_key, sub)
            
            if current_usage >= base_limit:
                return False, current_usage, base_limit, "You have reached your limit for this billing period."
                
            if increment:
                PlanLimitService._increment_usage(user, feature_key, sub)
                current_usage += 1
                
            return True, current_usage, base_limit, None

        return False, 0, 0, "Unknown feature type."

    @staticmethod
    def _get_current_usage(user, feature_key, subscription):
        period_start, period_end = PlanLimitService.get_billing_period(subscription)
        record = UsageRecord.objects.filter(
            user=user, 
            feature_key=feature_key, 
            period_start=period_start
        ).first()
        return record.current_usage if record else 0

    @staticmethod
    def _increment_usage(user, feature_key, subscription):
        period_start, period_end = PlanLimitService.get_billing_period(subscription)
        record, created = UsageRecord.objects.get_or_create(
            user=user,
            feature_key=feature_key,
            period_start=period_start,
            defaults={'period_end': period_end, 'current_usage': 0}
        )
        record.current_usage += 1
        record.save()
        return record.current_usage
