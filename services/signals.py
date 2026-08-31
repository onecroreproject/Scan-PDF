from django.contrib.auth.signals import user_logged_in, user_logged_out
from django.db.models.signals import post_save
from django.dispatch import receiver
from django.contrib.auth.models import User
from dynamic_qr.models import DynamicQRCode
from .models import Subscription, Payment
from .utils import log_activity

@receiver(user_logged_in)
def handle_user_login(sender, request, user, **kwargs):
    log_activity(user, "User Login", request)

@receiver(user_logged_out)
def handle_user_logout(sender, request, user, **kwargs):
    log_activity(user, "User Logout", request)

@receiver(post_save, sender=User)
def handle_user_registration(sender, instance, created, **kwargs):
    if created:
        log_activity(instance, "User Registration", status='Success')

@receiver(post_save, sender=DynamicQRCode)
def handle_qr_creation(sender, instance, created, **kwargs):
    if created:
        action_name = "Short URL Created" if instance.qr_type == 'custom-url' else "QR Code Created"
        log_activity(instance.user, f"{action_name} ({instance.qr_name})", status='Success')

@receiver(post_save, sender=Subscription)
def handle_subscription_changes(sender, instance, created, **kwargs):
    if created:
        log_activity(instance.user, f"Subscription Activated ({instance.plan.name})", status='Success')
        # Generate Entitlement Snapshot
        _create_subscription_snapshot(instance)
    else:
        if instance.status == 'Expired':
            log_activity(instance.user, f"Subscription Expired ({instance.plan.name})", status='Expired')
        elif instance.status == 'Cancelled':
            log_activity(instance.user, f"Subscription Cancelled ({instance.plan.name})", status='Cancelled')
        elif instance.status == 'Active' and not hasattr(instance, 'snapshot'):
            # In case it becomes active later without a snapshot
            _create_subscription_snapshot(instance)

def _create_subscription_snapshot(subscription):
    """Creates an immutable snapshot of features for a subscription."""
    from .models import SubscriptionSnapshot, PlanSectionFeature
    
    features_data = {}
    
    # 1. Fetch dynamic PlanSectionFeatures
    psf_list = PlanSectionFeature.objects.filter(
        section__plan=subscription.plan,
        section__is_enabled=True,
        is_enabled=True
    )
    
    for psf in psf_list:
        # Generate a slug key (e.g. 'Short URLs' -> 'short_urls')
        # This matches what limit_service expects or what we can use.
        # Fallbacks for exact key matches
        key_mapping = {
            'Short URLs': 'short_urls',
            'Dynamic QR Codes': 'dynamic_qrs',
            'Custom Alias': 'custom_alias',
            'Password Protection': 'password_protection',
            'Link Expiry': 'link_expiry',
            'GPS Tracking': 'gps_tracking',
            'Analytics': 'analytics',
        }
        key = key_mapping.get(psf.name, psf.name.lower().replace(' ', '_'))
        
        limit_val = None
        is_unlimited = psf.feature_type == 'UNLIMITED' or psf.is_unlimited
        if psf.feature_type == 'LIMIT':
            if subscription.billing_cycle == 'yearly':
                limit_val = psf.yearly_value
            else:
                limit_val = psf.monthly_value
                
        features_data[key] = {
            'enabled': True,
            'limit': limit_val,
            'is_unlimited': is_unlimited
        }
        
    # 2. Add fallback for legacy features if they exist and aren't in dynamic features yet
    legacy_keys = {
        'short_urls': subscription.plan.max_short_urls,
        'dynamic_qrs': subscription.plan.max_dynamic_qrs,
        'gps_tracking': getattr(subscription.plan, 'gps_tracking', False),
    }
    
    for lk, lval in legacy_keys.items():
        if lk not in features_data:
            if isinstance(lval, bool):
                features_data[lk] = {'enabled': lval, 'limit': None, 'is_unlimited': False}
            else:
                features_data[lk] = {'enabled': lval > 0, 'limit': lval, 'is_unlimited': False}
                
    SubscriptionSnapshot.objects.create(
        subscription=subscription,
        plan_name=subscription.plan.name,
        plan_code=subscription.plan.code,
        price_at_purchase=subscription.price_at_purchase,
        billing_cycle=subscription.billing_cycle,
        features_data=features_data
    )

@receiver(post_save, sender=Payment)
def handle_payment_status(sender, instance, created, **kwargs):
    if created:
        status = 'Success' if instance.payment_status == 'Paid' else 'Failed'
        log_activity(instance.user, f"Payment {status} (Amount: ₹{instance.amount}, TXN: {instance.transaction_id})", status=status)
