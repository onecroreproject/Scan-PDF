from django.shortcuts import redirect
from django.contrib import messages
from django.http import HttpResponseForbidden
from functools import wraps
from .models import Subscription, Plan
from dynamic_qr.models import DynamicQRCode


def get_user_subscription(user):
    """Retrieves or creates the user's active subscription."""
    if not user.is_authenticated:
        return None
    sub = Subscription.objects.filter(user=user, status='Active').select_related('plan').first()
    if not sub:
        # Fallback to Free subscription
        free_plan, _ = Plan.objects.get_or_create(code='free', defaults={'name': 'Free', 'is_default': True})
        sub = Subscription.objects.create(
            user=user,
            plan=free_plan,
            status='Active',
            billing_cycle='monthly',
            payment_status='Paid'
        )
    return sub


def check_dynamic_qr_limit(view_func):
    """Decorator to enforce active plan's dynamic QR generation limit via PlanFeatureService."""
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('dynamic_qr:login')

        from .plan_features import can_use_feature, get_feature_status
        # Dynamic QRs (non short URLs) use the qr_code feature
        if not can_use_feature(request.user, 'qr_code'):
            status = get_feature_status(request.user, 'qr_code')
            limit = status.get('limit', 'N/A')
            messages.error(
                request,
                f"You have reached your QR Code limit ({limit}) for this billing period. Please upgrade to create more."
            )
            return redirect('services:pricing')

        return view_func(request, *args, **kwargs)
    return _wrapped_view


def check_short_url_limit(view_func):
    """Decorator to enforce active plan's Short URL generation limit via PlanFeatureService."""
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('dynamic_qr:login')

        from .plan_features import can_use_feature, get_feature_status
        if not can_use_feature(request.user, 'qr_code'):
            status = get_feature_status(request.user, 'qr_code')
            limit = status.get('limit', 'N/A')
            messages.error(
                request,
                f"You have reached your Short URL limit ({limit}) for this billing period. Please upgrade to create more."
            )
            return redirect('services:pricing')

        return view_func(request, *args, **kwargs)
    return _wrapped_view


def require_premium_feature(feature_code):
    """Decorator to enforce boolean plan feature flags via PlanFeatureService."""
    def decorator(view_func):
        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect('dynamic_qr:login')

            from .plan_features import has_feature, get_user_plan
            if not has_feature(request.user, feature_code):
                plan = get_user_plan(request.user)
                plan_name = plan.name if plan else 'your current'
                feature_display = feature_code.replace('_', ' ').title()
                messages.error(
                    request,
                    f"The feature '{feature_display}' is not available on your {plan_name} plan. Please upgrade to access this feature."
                )
                return redirect('services:pricing')

            return view_func(request, *args, **kwargs)
        return _wrapped_view
    return decorator
