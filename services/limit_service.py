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
        """Checks if the user's active plan allows a specific feature using the Entitlement Snapshot."""
        sub = PlanLimitService.get_active_subscription(user)
        if not sub:
            return False
            
        try:
            snapshot = sub.snapshot
        except Exception:
            return False

        feature_data = snapshot.features_data.get(feature_key)
        if not feature_data:
            return False
            
        if not feature_data.get('enabled', False):
            return False
            
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
        Checks if the user has reached their limit for the given feature using Entitlement Snapshot.
        If increment is True, increments the usage counter atomically.
        Returns a tuple: (allowed: bool, current_usage: int, limit: int/str, error_msg: str)
        """
        sub = PlanLimitService.get_active_subscription(user)
        if not sub:
            return False, 0, 0, "No active subscription found."
            
        try:
            snapshot = sub.snapshot
        except Exception:
            return False, 0, 0, "No entitlement snapshot found for this subscription."
            
        feature_data = snapshot.features_data.get(feature_key)
        if not feature_data or not feature_data.get('enabled', False):
            return False, 0, 0, "Feature not available in your current plan."
            
        # Check unlimited
        if feature_data.get('is_unlimited', False):
            if increment:
                PlanLimitService._increment_usage(user, feature_key, sub)
            current = PlanLimitService._get_current_usage(user, feature_key, sub)
            return True, current, 'Unlimited', None
            
        # Check numeric limit
        base_limit = feature_data.get('limit')
        if base_limit is None:
            # If it's just a boolean flag with no limit concept, it's unlimited
            return True, 0, 'Unlimited', None

        # Apply overrides
        override = UsageOverride.objects.filter(user=user, feature_key=feature_key).filter(
            models.Q(expires_at__isnull=True) | models.Q(expires_at__gt=timezone.now())
        ).first()

        if override:
            if override.override_limit is not None:
                base_limit = override.override_limit
            else:
                base_limit += override.additional_allowance

        if increment:
            # Atomic increment
            new_usage = PlanLimitService._increment_usage_atomic(user, feature_key, sub, base_limit)
            if new_usage == -1:
                return False, base_limit, base_limit, "You have reached your limit for this billing period."
            return True, new_usage, base_limit, None
        else:
            current_usage = PlanLimitService._get_current_usage(user, feature_key, sub)
            if current_usage >= base_limit:
                return False, current_usage, base_limit, "You have reached your limit for this billing period."
            return True, current_usage, base_limit, None

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
    def _increment_usage_atomic(user, feature_key, subscription, limit):
        """Atomically increment usage. Returns new usage, or -1 if limit reached."""
        from django.db import transaction
        period_start, period_end = PlanLimitService.get_billing_period(subscription)
        
        with transaction.atomic():
            record, created = UsageRecord.objects.select_for_update().get_or_create(
                user=user,
                feature_key=feature_key,
                period_start=period_start,
                defaults={'period_end': period_end, 'current_usage': 0}
            )
            if record.current_usage >= limit:
                return -1
            record.current_usage += 1
            record.save()
            return record.current_usage

    @staticmethod
    def _increment_usage(user, feature_key, subscription):
        """Non-atomic legacy increment for unlimited features."""
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
