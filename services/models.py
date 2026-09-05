from django.db import models
from django.contrib.auth.models import User
from django.utils import timezone
import datetime

class Plan(models.Model):
    name = models.CharField(max_length=50, unique=True) # FREE, PRO, BUSINESS+
    code = models.CharField(max_length=20, unique=True) # free, pro, business_plus
    monthly_price = models.IntegerField(default=0) # INR
    yearly_price = models.IntegerField(default=0) # INR

    pricing_type = models.CharField(
        max_length=20,
        choices=[('fixed', 'Fixed Price'), ('contact', 'Contact Us')],
        default='fixed'
    )

    description = models.TextField(blank=True, help_text="Plan description")
    is_active = models.BooleanField(default=True)
    is_popular = models.BooleanField(default=False)
    display_order = models.IntegerField(default=0)
    setup_fee = models.IntegerField(default=0, help_text="Optional setup fee in INR")
    is_default = models.BooleanField(default=False, help_text="Is this the default plan for new users?")

    # Legacy Limits (kept for backward compat — runtime checks use PlanFeature)
    max_dynamic_qrs = models.IntegerField(default=5)
    max_short_urls = models.IntegerField(default=10)
    max_team_members = models.IntegerField(default=0)
    max_domains = models.IntegerField(default=0)
    max_api_requests = models.IntegerField(default=0)

    # Legacy boolean plan flags (deprecated — use PlanFeature)
    analytics_access = models.BooleanField(default=False)
    csv_export = models.BooleanField(default=False)
    gps_tracking = models.BooleanField(default=False)
    bulk_qr = models.BooleanField(default=False)
    webhooks = models.BooleanField(default=False)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.name


# ─────────────────────────────────────────────────────────────────────────────
# Feature Master — the canonical list of all Short URL features
# ─────────────────────────────────────────────────────────────────────────────

FEATURE_CODES = [
    'header',
    'qr_code',
    'password_protection',
    'link_expiry',
    'gps_tracking',
    'analytics',
    'custom_alias',
    'csv_export',
    'pdf_report',
]

class Feature(models.Model):
    FEATURE_TYPES = [
        ('BOOLEAN', 'Boolean (ON/OFF)'),
        ('NUMERIC', 'Numeric Limit'),
        ('UNLIMITED', 'Unlimited'),
        ('SELECT', 'Select (Text/Choice)'),
        ('DURATION', 'Duration (Text)'),
    ]
    # 'key' kept for backward compat; 'code' is the canonical field used by the new service
    key = models.CharField(max_length=50, unique=True, help_text="e.g. header, qr_code, analytics")
    name = models.CharField(max_length=100)
    type = models.CharField(max_length=20, choices=FEATURE_TYPES, default='BOOLEAN')
    description = models.TextField(blank=True)
    section = models.CharField(max_length=100, default='SHORT URL', help_text="Section this feature belongs to")
    is_public = models.BooleanField(default=True, help_text="Show on public pricing page")
    is_active = models.BooleanField(default=True)
    display_order = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    # 'code' property alias for 'key' (new code should use feature.code)
    @property
    def code(self):
        return self.key

    class Meta:
        ordering = ['display_order', 'name']

    def __str__(self):
        return f"{self.name} ({self.key})"


# ─────────────────────────────────────────────────────────────────────────────
# PlanFeature — the SINGLE SOURCE OF TRUTH for runtime entitlement
# ─────────────────────────────────────────────────────────────────────────────

class PlanFeature(models.Model):
    """
    Links a Plan to a Feature with its specific limit configuration.
    This is the SINGLE SOURCE OF TRUTH for runtime feature entitlement.
    Admin changes here take effect immediately — no restart required.
    """
    plan = models.ForeignKey(Plan, on_delete=models.CASCADE, related_name='plan_features')
    feature = models.ForeignKey(Feature, on_delete=models.CASCADE, related_name='plan_features')

    # Entitlement fields
    enabled = models.BooleanField(default=False, help_text="Whether this feature is enabled for this plan")
    monthly_limit = models.IntegerField(null=True, blank=True, help_text="Monthly usage limit")
    yearly_limit = models.IntegerField(null=True, blank=True, help_text="Yearly usage limit")
    is_unlimited = models.BooleanField(default=False, help_text="If True, ignore limit entirely for all billing cycles")
    # Analytics-specific: how many days of history the user can access
    history_days = models.IntegerField(
        null=True, blank=True,
        help_text="For analytics feature: how many days of history are accessible"
    )

    # Legacy value fields (kept for backward compat with old code)
    value_boolean = models.BooleanField(default=False)
    value_numeric = models.IntegerField(null=True, blank=True)
    value_text = models.CharField(max_length=255, null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ('plan', 'feature')

    def __str__(self):
        return f"{self.plan.name} — {self.feature.name}"

    def get_display_value(self):
        if not self.enabled:
            return "OFF"
        if self.is_unlimited:
            return "Unlimited"
        if self.feature.key == 'analytics':
            return f"{self.history_days or 0} days"
        
        m_limit = self.monthly_limit if self.monthly_limit is not None else 0
        y_limit = self.yearly_limit if self.yearly_limit is not None else 0
        return f"{m_limit}/mo · {y_limit}/yr"


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


class SubscriptionSnapshot(models.Model):
    """
    Immutable snapshot of plan entitlements at subscription time.
    DEPRECATED for runtime checks — use PlanFeatureService (services/plan_features.py) instead.
    Kept for historical audit purposes only.
    """
    subscription = models.OneToOneField(Subscription, on_delete=models.CASCADE, related_name='snapshot')
    plan_name = models.CharField(max_length=100)
    plan_code = models.CharField(max_length=50)
    price_at_purchase = models.IntegerField(default=0)
    billing_cycle = models.CharField(max_length=20)
    features_data = models.JSONField(
        default=dict,
        help_text="Serialized dict of feature configurations at subscription time (audit only)."
    )
    snapshot_date = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Snapshot for {self.subscription}"


class ActivityLog(models.Model):
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='activity_logs')
    action = models.CharField(max_length=255)
    created_at = models.DateTimeField(auto_now_add=True)
    ip_address = models.GenericIPAddressField(null=True, blank=True)
    status = models.CharField(max_length=50, default='Success')

    def __str__(self):
        return f"{self.user.username if self.user else 'System'} - {self.action} ({self.created_at})"


class UsageRecord(models.Model):
    """Tracks per-user, per-feature, per-billing-period usage counts."""
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
    """Per-user limit override set by admin (e.g. bonus allowance)."""
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='usage_overrides')
    feature_key = models.CharField(max_length=50)
    additional_allowance = models.IntegerField(default=0)
    override_limit = models.IntegerField(null=True, blank=True, help_text="Replaces the plan limit if set")
    expires_at = models.DateTimeField(null=True, blank=True)
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='+')
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Override for {self.user.username} on {self.feature_key}"


# ─────────────────────────────────────────────────────────────────────────────
# Dynamic Plan Builder — Plan-local Sections & Features (VISUAL DISPLAY ONLY)
# These are for the pricing page display. Runtime checks use PlanFeature above.
# ─────────────────────────────────────────────────────────────────────────────

class PlanSection(models.Model):
    """A section header inside a pricing plan (e.g. 'SHORT URL'). Display only."""
    plan = models.ForeignKey(Plan, on_delete=models.CASCADE, related_name='sections')
    name = models.CharField(max_length=100, help_text='Section header shown on pricing card')
    is_enabled = models.BooleanField(default=True)
    display_order = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['display_order', 'id']

    def __str__(self):
        return f"{self.plan.name} → {self.name}"


class PlanSectionFeature(models.Model):
    """A single feature row inside a PlanSection for visual display on pricing page."""

    FEATURE_TYPE_CHOICES = [
        ('BOOLEAN', 'Boolean (ON/OFF — shown as ✓ Feature Name)'),
        ('LIMIT', 'Limit (numeric value + label, e.g. 100 Short URLs / month)'),
        ('TEXT', 'Text (custom label, e.g. 30 Days Analytics History)'),
        ('UNLIMITED', 'Unlimited (shown as ✓ Unlimited Feature Name)'),
    ]

    section = models.ForeignKey(PlanSection, on_delete=models.CASCADE, related_name='features')
    name = models.CharField(max_length=200, help_text='Feature name, e.g. Short URLs')
    description = models.TextField(blank=True, help_text='Optional internal note')
    feature_type = models.CharField(max_length=20, choices=FEATURE_TYPE_CHOICES, default='BOOLEAN')

    monthly_value = models.IntegerField(null=True, blank=True, help_text='Numeric value for monthly billing')
    yearly_value = models.IntegerField(null=True, blank=True, help_text='Numeric value for yearly billing')
    monthly_label = models.CharField(max_length=255, blank=True)
    yearly_label = models.CharField(max_length=255, blank=True)
    monthly_text = models.CharField(max_length=255, blank=True)
    yearly_text = models.CharField(max_length=255, blank=True)
    is_unlimited = models.BooleanField(default=False)
    is_enabled = models.BooleanField(default=True)
    display_order = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['display_order', 'id']

    def __str__(self):
        return f"{self.section.plan.name} / {self.section.name} → {self.name}"

    def get_monthly_display(self):
        if self.feature_type == 'UNLIMITED' or self.is_unlimited:
            return f"Unlimited {self.name}"
        if self.feature_type == 'BOOLEAN':
            return self.name
        if self.feature_type == 'LIMIT':
            if self.monthly_value is not None:
                label = self.monthly_label or self.name
                return f"{self.monthly_value} {label}"
            return self.name
        if self.feature_type == 'TEXT':
            text = self.monthly_text or ''
            label = self.monthly_label or self.name
            return f"{text} {label}".strip() if text else self.name
        return self.name

    def get_yearly_display(self):
        if self.feature_type == 'UNLIMITED' or self.is_unlimited:
            return f"Unlimited {self.name}"
        if self.feature_type == 'BOOLEAN':
            return self.name
        if self.feature_type == 'LIMIT':
            if self.yearly_value is not None:
                label = self.yearly_label or self.name
                return f"{self.yearly_value} {label}"
            return self.name
        if self.feature_type == 'TEXT':
            text = self.yearly_text or ''
            label = self.yearly_label or self.name
            return f"{text} {label}".strip() if text else self.name
        return self.name
