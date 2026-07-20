from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import datetime

class Plan(models.Model):
    name = models.CharField(max_length=50, unique=True) # FREE, PRO, BUSINESS, BUSINESS+
    code = models.CharField(max_length=20, unique=True) # free, pro, business, business_plus
    monthly_price = models.IntegerField(default=0) # INR
    yearly_price = models.IntegerField(default=0) # INR
    
    # Limits and features
    max_dynamic_qrs = models.IntegerField(default=5)
    max_short_urls = models.IntegerField(default=10)
    max_team_members = models.IntegerField(default=0)
    max_domains = models.IntegerField(default=0)
    max_api_requests = models.IntegerField(default=0)
    
    analytics_access = models.BooleanField(default=False)
    csv_export = models.BooleanField(default=False)
    gps_tracking = models.BooleanField(default=False)
    bulk_qr = models.BooleanField(default=False)
    webhooks = models.BooleanField(default=False)
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name

class Subscription(models.Model):
    STATUS_CHOICES = [
        ('Active', 'Active'),
        ('Expired', 'Expired'),
        ('Cancelled', 'Cancelled'),
        ('Pending', 'Pending'),
    ]
    
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='subscriptions')
    plan = models.ForeignKey(Plan, on_delete=models.CASCADE, related_name='subscriptions')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='Pending')
    start_date = models.DateTimeField(default=timezone.now)
    end_date = models.DateTimeField(null=True, blank=True) # Null for FREE
    billing_cycle = models.CharField(max_length=20, choices=[('monthly', 'Monthly'), ('yearly', 'Yearly')], default='monthly')
    payment_status = models.CharField(max_length=20, default='Pending')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.user.username} - {self.plan.name} ({self.status})"

class Payment(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='payments')
    subscription = models.ForeignKey(Subscription, on_delete=models.SET_NULL, null=True, blank=True, related_name='payments')
    amount = models.IntegerField() # In INR
    currency = models.CharField(max_length=10, default='INR')
    transaction_id = models.CharField(max_length=100, unique=True)
    payment_status = models.CharField(max_length=20, default='Pending')
    gateway = models.CharField(max_length=50, default='Simulator')
    payment_mode = models.CharField(max_length=50, default='Card')
    receipt_number = models.CharField(max_length=100, blank=True, null=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.transaction_id} - {self.user.username} (₹{self.amount})"


class ActivityLog(models.Model):
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='activity_logs')
    action = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    status = models.CharField(max_length=50, default='Success')

    def __str__(self):
        return f"{self.user.username if self.user else 'System'} - {self.action} ({self.created_at})"
