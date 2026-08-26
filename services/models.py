from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import datetime

class Plan(models.Model):
    name = models.CharField(max_length=50, unique=True) # FREE, PRO, BUSINESS, BUSINESS+
    code = models.CharField(max_length=20, unique=True) # free, pro, business, business_plus
    monthly_price = models.IntegerField(default=0) # INR
    yearly_price = models.IntegerField(default=0) # INR
    
    pricing_type = models.CharField(max_length=20, choices=[('fixed', 'Fixed Price'), ('contact', 'Contact Us')], default='fixed')
    
    description = models.TextField(blank=True, help_text="Plan description")
    is_active = models.BooleanField(default=True)
    is_popular = models.BooleanField(default=False)
    display_order = models.IntegerField(default=0)
    setup_fee = models.IntegerField(default=0, help_text="Optional setup fee in INR")
    is_default = models.BooleanField(default=False, help_text="Is this the default plan for new users?")
    
    # Legacy Limits (Kept for backward compatibility but plan to migrate to Feature system)
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

class Feature(models.Model):
    FEATURE_TYPES = [
        ('BOOLEAN', 'Boolean (ON/OFF)'),
        ('NUMERIC', 'Numeric Limit'),
        ('UNLIMITED', 'Unlimited'),
        ('SELECT', 'Select (Text/Choice)'),
        ('DURATION', 'Duration (Text)'),
    ]
    key = models.CharField(max_length=50, unique=True, help_text="e.g. short_urls, gps_tracking")
    name = models.CharField(max_length=100)
    type = models.CharField(max_length=20, choices=FEATURE_TYPES, default='BOOLEAN')
    description = models.TextField(blank=True)
    is_public = models.BooleanField(default=True, help_text="Show on public pricing page")
    display_order = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['display_order', 'name']

    def __str__(self):
        return f"{self.name} ({self.key})"

class PlanFeature(models.Model):
    plan = models.ForeignKey(Plan, on_delete=models.CASCADE, related_name='features')
    feature = models.ForeignKey(Feature, on_delete=models.CASCADE, related_name='plan_features')
    
    value_boolean = models.BooleanField(default=False)
    value_numeric = models.IntegerField(null=True, blank=True)
    value_text = models.CharField(max_length=255, null=True, blank=True)
    
    class Meta:
        unique_together = ('plan', 'feature')

    def __str__(self):
        return f"{self.plan.name} - {self.feature.name}"

    def get_display_value(self):
        if self.feature.type == 'BOOLEAN':
            return "ON" if self.value_boolean else "OFF"
        elif self.feature.type == 'UNLIMITED':
            return "Unlimited"
        elif self.feature.type == 'NUMERIC':
            return str(self.value_numeric)
        return self.value_text or ""

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
    
    # Grandfathering & Billing Details
    price_at_purchase = models.IntegerField(default=0, help_text="Price stored at time of purchase")
    currency = models.CharField(max_length=10, default='INR')
    auto_renew = models.BooleanField(default=True)
    renewal_date = models.DateTimeField(null=True, blank=True)
    cancelled_at = models.DateTimeField(null=True, blank=True)
    
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

class UsageRecord(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='usage_records')
    feature_key = models.CharField(max_length=50, db_index=True)
    period_start = models.DateTimeField()
    period_end = models.DateTimeField()
    current_usage = models.IntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ('user', 'feature_key', 'period_start')

    def __str__(self):
        return f"{self.user.username} - {self.feature_key}: {self.current_usage}"

class UsageOverride(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='usage_overrides')
    feature_key = models.CharField(max_length=50)
    additional_allowance = models.IntegerField(default=0)
    override_limit = models.IntegerField(null=True, blank=True, help_text="Replaces the plan limit if set")
    expires_at = models.DateTimeField(null=True, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='+')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Override for {self.user.username} on {self.feature_key}"
