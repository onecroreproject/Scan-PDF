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
    else:
        if instance.status == 'Expired':
            log_activity(instance.user, f"Subscription Expired ({instance.plan.name})", status='Expired')
        elif instance.status == 'Cancelled':
            log_activity(instance.user, f"Subscription Cancelled ({instance.plan.name})", status='Cancelled')

@receiver(post_save, sender=Payment)
def handle_payment_status(sender, instance, created, **kwargs):
    if created:
        status = 'Success' if instance.payment_status == 'Paid' else 'Failed'
        log_activity(instance.user, f"Payment {status} (Amount: ₹{instance.amount}, TXN: {instance.transaction_id})", status=status)
