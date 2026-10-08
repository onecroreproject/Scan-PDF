"""
services/plan_features.py
=========================
CENTRALIZED FEATURE ENTITLEMENT SERVICE

This is the SINGLE SOURCE OF TRUTH for runtime feature permission checks.
All Short URL backend views must call this service — never check plan.name directly.

Key design principles:
- Always reads LIVE PlanFeature records (no stale snapshots)
- Admin changes take effect immediately after save
- Atomic usage increments prevent race conditions
- Delta-based usage: only counts NEW feature activations, not edits
- N+1 avoided via get_all_feature_statuses() batch method
"""

from django.utils import timezone
from django.db import transaction, models as django_models
import datetime

from .models import (
    Plan, Subscription, PlanFeature, Feature,
    UsageRecord, UsageOverride, FEATURE_CODES
)


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _get_active_subscription(user):
    """Returns the active subscription for a user, or creates a FREE one."""
    if not user or not user.is_authenticated:
        return None
    sub = Subscription.objects.filter(user=user, status='Active').select_related('plan').first()
    if not sub:
        try:
            free_plan = Plan.objects.get(code='free')
        except Plan.DoesNotExist:
            return None
        sub = Subscription.objects.create(
            user=user,
            plan=free_plan,
            status='Active',
            billing_cycle='monthly',
            payment_status='Paid',
        )
    return sub


def _get_billing_period(subscription):
    """Returns (period_start, period_end) for the current billing cycle."""
    now = timezone.now()
    start_date = subscription.start_date
    from dateutil.relativedelta import relativedelta

    if subscription.billing_cycle == 'yearly':
        years_diff = relativedelta(now, start_date).years
        period_start = start_date + relativedelta(years=years_diff)
        period_end = period_start + relativedelta(years=1)
    else:
        # Monthly: calculate exact month boundary using relativedelta
        diff = relativedelta(now, start_date)
        months_diff = diff.years * 12 + diff.months
        period_start = start_date + relativedelta(months=months_diff)
        period_end = period_start + relativedelta(months=1)

    return period_start, period_end


def _get_current_usage(user, feature_key, subscription):
    """Returns the current usage count for the billing period."""
    period_start, _ = _get_billing_period(subscription)
    record = UsageRecord.objects.filter(
        user=user,
        feature_key=feature_key,
        period_start__lte=timezone.now(),
        period_end__gte=timezone.now(),
    ).order_by('-period_start').first()
    # Fallback: exact period_start match
    if not record:
        record = UsageRecord.objects.filter(
            user=user,
            feature_key=feature_key,
            period_start=period_start,
        ).first()
    return record.current_usage if record else 0


def _get_effective_limit(user, feature_key, base_limit):
    """Applies UsageOverride to base_limit and returns effective limit.
    Handles multiple overrides deterministically:
    - Absolute limits (override_limit) take precedence, using the MAXIMUM value.
    - If no absolute limits exist, sums all active additional allowances.
    """
    overrides = UsageOverride.objects.filter(
        user=user,
        feature_key=feature_key,
    ).filter(
        django_models.Q(expires_at__isnull=True) | django_models.Q(expires_at__gt=timezone.now())
    )

    absolute_limits = [o.override_limit for o in overrides if o.override_limit is not None]
    if absolute_limits:
        return max(absolute_limits)

    total_additional = sum(o.additional_allowance for o in overrides)
    return base_limit + total_additional


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def get_user_plan(user):
    """Returns the active Plan for the user."""
    sub = _get_active_subscription(user)
    return sub.plan if sub else None


def get_plan_feature(user, feature_code):
    """
    Returns the PlanFeature for (user's plan, feature_code) or None.
    feature_code is the Feature.key value (e.g. 'qr_code', 'analytics').
    """
    plan = get_user_plan(user)
    if not plan:
        return None
    return PlanFeature.objects.filter(
        plan=plan,
        feature__key=feature_code,
    ).select_related('feature').first()


def has_feature(user, feature_code):
    """Returns True if the user's plan has the feature enabled and the core feature is active."""
    pf = get_plan_feature(user, feature_code)
    return bool(pf and pf.enabled and pf.feature.is_active)


def get_feature_limit(user, feature_code):
    """
    Returns the effective numeric limit for the feature, or None if unlimited.
    Returns None for analytics feature (use history_days instead).
    """
    pf = get_plan_feature(user, feature_code)
    if not pf or not pf.enabled or not pf.feature.is_active:
        return 0
    if pf.is_unlimited:
        return None  # None = unlimited
    sub = _get_active_subscription(user)
    base_limit = pf.yearly_limit if sub and sub.billing_cycle == 'yearly' else pf.monthly_limit
    return _get_effective_limit(user, feature_code, base_limit or 0)


def get_feature_usage(user, feature_code):
    """Returns current period usage count for the feature."""
    sub = _get_active_subscription(user)
    if not sub:
        return 0
    return _get_current_usage(user, feature_code, sub)


def get_feature_remaining(user, feature_code):
    """Returns remaining usage (None if unlimited, 0 if disabled)."""
    pf = get_plan_feature(user, feature_code)
    if not pf or not pf.enabled or not pf.feature.is_active:
        return 0
    if pf.is_unlimited:
        return None
    limit = get_feature_limit(user, feature_code)
    used = get_feature_usage(user, feature_code)
    return max(0, (limit or 0) - used)


def can_use_feature(user, feature_code):
    """Returns True if the user can use the feature right now (enabled + not at limit)."""
    pf = get_plan_feature(user, feature_code)
    if not pf or not pf.enabled or not pf.feature.is_active:
        return False
    if pf.is_unlimited:
        return True
    remaining = get_feature_remaining(user, feature_code)
    return remaining is None or remaining > 0


def get_feature_status(user, feature_code):
    """
    Returns a full status dict for a single feature.

    Shape:
    {
        "enabled": bool,
        "unlimited": bool,
        "limit": int or None,
        "used": int,
        "remaining": int or None,
        "limit_reached": bool,
        "history_days": int or None,  # analytics only
    }
    """
    pf = get_plan_feature(user, feature_code)
    sub = _get_active_subscription(user)
    cycle = sub.billing_cycle if sub else 'monthly'
    
    if not pf or not pf.enabled:
        return {
            "enabled": False,
            "billing_cycle": cycle,
            "unlimited": False,
            "limit": 0,
            "used": 0,
            "remaining": 0,
            "limit_reached": True,
            "history_days": None,
        }

    used = _get_current_usage(user, feature_code, sub) if sub else 0

    if pf.is_unlimited:
        return {
            "enabled": True,
            "billing_cycle": cycle,
            "unlimited": True,
            "limit": None,
            "used": used,
            "remaining": None,
            "limit_reached": False,
            "history_days": pf.history_days,
        }

    base_limit = pf.yearly_limit if cycle == 'yearly' else pf.monthly_limit
    effective_limit = _get_effective_limit(user, feature_code, base_limit or 0)
    remaining = max(0, effective_limit - used)
    return {
        "enabled": True,
        "billing_cycle": cycle,
        "unlimited": False,
        "limit": effective_limit,
        "used": used,
        "remaining": remaining,
        "limit_reached": remaining <= 0,
        "history_days": pf.history_days,
    }


def get_all_feature_statuses(user):
    """
    Returns a dict of feature_code → status for ALL 9 Short URL features.
    Uses a single DB query to avoid N+1.

    Returns:
    {
        "header": { "enabled": ..., "unlimited": ..., "limit": ..., "used": ..., ... },
        "qr_code": { ... },
        ...
    }
    """
    plan = get_user_plan(user)
    sub = _get_active_subscription(user)

    # Single query: all PlanFeatures for this plan + all 9 feature codes
    plan_features = {}
    if plan:
        pf_qs = PlanFeature.objects.filter(
            plan=plan,
            feature__key__in=FEATURE_CODES,
        ).select_related('feature')
        for pf in pf_qs:
            plan_features[pf.feature.key] = pf

    # Single query: all usage records for current period
    usage_map = {}
    if sub:
        period_start, _ = _get_billing_period(sub)
        records = UsageRecord.objects.filter(
            user=user,
            feature_key__in=FEATURE_CODES,
        ).filter(
            django_models.Q(period_start__lte=timezone.now()) &
            django_models.Q(period_end__gte=timezone.now())
        )
        for rec in records:
            usage_map[rec.feature_key] = rec.current_usage

    result = {}
    cycle = sub.billing_cycle if sub else 'monthly'
    
    for code in FEATURE_CODES:
        pf = plan_features.get(code)
        used = usage_map.get(code, 0)

        if not pf or not pf.enabled:
            result[code] = {
                "enabled": False,
                "billing_cycle": cycle,
                "unlimited": False,
                "limit": 0,
                "used": used,
                "remaining": 0,
                "limit_reached": True,
                "history_days": None,
            }
            continue

        if pf.is_unlimited:
            result[code] = {
                "enabled": True,
                "billing_cycle": cycle,
                "unlimited": True,
                "limit": None,
                "used": used,
                "remaining": None,
                "limit_reached": False,
                "history_days": pf.history_days,
            }
            continue

        base_limit = pf.yearly_limit if cycle == 'yearly' else pf.monthly_limit
        effective_limit = _get_effective_limit(user, code, base_limit or 0)
        remaining = max(0, effective_limit - used)
        result[code] = {
            "enabled": True,
            "billing_cycle": cycle,
            "unlimited": False,
            "limit": effective_limit,
            "used": used,
            "remaining": remaining,
            "limit_reached": remaining <= 0,
            "history_days": pf.history_days,
        }

    return result


def increment_feature_usage(user, feature_code):
    """
    Atomically increments usage for the feature.
    Returns new usage count, or -1 if limit was already reached.
    For unlimited features: increments and returns new count (for tracking).
    """
    pf = get_plan_feature(user, feature_code)
    if not pf or not pf.enabled:
        return -1

    sub = _get_active_subscription(user)
    if not sub:
        return -1

    period_start, period_end = _get_billing_period(sub)

    if pf.is_unlimited:
        # Still track usage for analytics, but never block
        with transaction.atomic():
            record, _ = UsageRecord.objects.select_for_update().get_or_create(
                user=user,
                feature_key=feature_code,
                period_start=period_start,
                defaults={'period_end': period_end, 'current_usage': 0}
            )
            record.current_usage += 1
            record.save(update_fields=['current_usage', 'updated_at'])
            return record.current_usage

    base_limit = pf.yearly_limit if sub.billing_cycle == 'yearly' else pf.monthly_limit
    effective_limit = _get_effective_limit(user, feature_code, base_limit or 0)

    with transaction.atomic():
        record, _ = UsageRecord.objects.select_for_update().get_or_create(
            user=user,
            feature_key=feature_code,
            period_start=period_start,
            defaults={'period_end': period_end, 'current_usage': 0}
        )
        if record.current_usage >= effective_limit:
            return -1  # Limit already reached
        record.current_usage += 1
        record.save(update_fields=['current_usage', 'updated_at'])
        return record.current_usage


# ─────────────────────────────────────────────────────────────────────────────
# Delta-based helpers for Short URL create/edit
# ─────────────────────────────────────────────────────────────────────────────

def check_and_increment_short_url_features(user, new_state, existing_qr=None):
    """
    Checks and atomically increments usage for all activated features in a
    Short URL create/edit operation.

    new_state: dict of booleans indicating which features are being used:
        {
            'header': bool,       # True if header_value is non-empty
            'qr_code': bool,
            'password_protection': bool,
            'link_expiry': bool,
            'gps_tracking': bool,
            'custom_alias': bool,
            'shorturl_utm': bool,
            'shorturl_cloaking': bool,
            'short_url': bool, # True on creation/edit requests
        }

    existing_qr: DynamicQRCode instance (None for new creation)

    Returns:
        (ok: bool, error_code: str|None, error_message: str|None)
        error_code: 'feature_not_available' | 'feature_limit_reached'
    """
    subscription = _get_active_subscription(user)
    if not subscription:
        return (False, 'feature_not_available', 'No active subscription was found.')

    # Serialize all feature activations for this user so concurrent requests
    # cannot both consume the final credit.
    with transaction.atomic():
        Subscription.objects.select_for_update().get(pk=subscription.pk)
        features_to_increment = []
        feature_check_map = {
            'header': lambda qr: bool(qr and qr.header),
            'qr_code': lambda qr: bool(qr and qr.qr_enabled),
            'password_protection': lambda qr: bool(qr and qr.password),
            'link_expiry': lambda qr: bool(qr and qr.expiry_date),
            'gps_tracking': lambda qr: bool(qr and qr.require_gps),
            'custom_alias': lambda qr: bool(qr and qr.custom_alias),
            'shorturl_utm': lambda qr: bool(qr and qr.utm_enabled),
            'shorturl_cloaking': lambda qr: bool(qr and qr.cloaking_enabled),
            'short_url': lambda qr: qr is not None,  # Base quota: only consumed on initial creation
        }
        feature_display_names = {
            'header': 'Custom Header',
            'qr_code': 'QR Code',
            'password_protection': 'Password Protection',
            'link_expiry': 'Link Expiry',
            'gps_tracking': 'GPS Tracking',
            'custom_alias': 'Custom Alias',
            'shorturl_utm': 'UTM Parameters',
            'shorturl_cloaking': 'URL Cloaking',
            'short_url': 'Short URLs',
        }

        for code, is_active_now in new_state.items():
            if code not in feature_check_map:
                continue
            if is_active_now and not feature_check_map[code](existing_qr):
                if not has_feature(user, code):
                    return (
                        False,
                        'feature_not_available',
                        f"{feature_display_names.get(code, code)} is not available in your current plan."
                    )
                if not can_use_feature(user, code):
                    status = get_feature_status(user, code)
                    remaining_msg = '' if status['unlimited'] else f" ({status['used']}/{status['limit']} used)"
                    return (
                        False,
                        'feature_limit_reached',
                        f"You have reached your {feature_display_names.get(code, code)} limit for this billing period{remaining_msg}."
                    )
                features_to_increment.append(code)

        for code in features_to_increment:
            result = increment_feature_usage(user, code)
            if result == -1:
                return (
                    False,
                    'feature_limit_reached',
                    f"You have reached your {feature_display_names.get(code, code)} limit. Please try again."
                )

    return (True, None, None)
