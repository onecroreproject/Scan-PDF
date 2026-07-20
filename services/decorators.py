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
    sub = Subscription.objects.filter(user=user, status='Active').first()
    if not sub:
        # Fallback to Free subscription
        free_plan = Plan.objects.get(code='free')
        sub = Subscription.objects.create(
            user=user,
            plan=free_plan,
            status='Active',
            billing_cycle='monthly',
            payment_status='Paid'
        )
    return sub

def check_dynamic_qr_limit(view_func):
    """Decorator to enforce active plan's dynamic QR generation limit."""
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('dynamic_qr:login')
            
        sub = get_user_subscription(request.user)
        plan = sub.plan
        
        # Count non-short URLs (standard Dynamic QRs)
        qr_count = DynamicQRCode.objects.filter(user=request.user).exclude(qr_type='custom-url').count()
        
        if qr_count >= plan.max_dynamic_qrs:
            messages.error(
                request, 
                f"You have reached the maximum Dynamic QR limit ({plan.max_dynamic_qrs}) of your {plan.name} subscription. Please upgrade to create more."
            )
            return redirect('services:pricing')
            
        return view_func(request, *args, **kwargs)
    return _wrapped_view

def check_short_url_limit(view_func):
    """Decorator to enforce active plan's Short URL generation limit."""
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect('dynamic_qr:login')
            
        sub = get_user_subscription(request.user)
        plan = sub.plan
        
        # Count short URLs
        short_count = DynamicQRCode.objects.filter(user=request.user, qr_type='custom-url').count()
        
        if short_count >= plan.max_short_urls:
            messages.error(
                request, 
                f"You have reached the maximum Short URL limit ({plan.max_short_urls}) of your {plan.name} subscription. Please upgrade to create more."
            )
            return redirect('services:pricing')
            
        return view_func(request, *args, **kwargs)
    return _wrapped_view

def require_premium_feature(feature_name):
    """Decorator to enforce boolean plan configuration flags (e.g. analytics, bulk_qr, webhooks)."""
    def decorator(view_func):
        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect('dynamic_qr:login')
                
            sub = get_user_subscription(request.user)
            plan = sub.plan
            
            if not getattr(plan, feature_name, False):
                messages.error(
                    request, 
                    f"The feature '{feature_name.replace('_', ' ').title()}' is not available on your {plan.name} plan. Please upgrade to access this feature."
                )
                return redirect('services:pricing')
                
            return view_func(request, *args, **kwargs)
        return _wrapped_view
    return decorator
