from django.contrib import admin
from django.http import HttpResponse
from django.utils.safestring import mark_safe
from django.utils import timezone
from django.db.models import Sum, Q
from django.contrib.auth.models import User
from django.contrib.auth.admin import UserAdmin
from .models import Plan, Subscription, Payment, ActivityLog
from dynamic_qr.models import DynamicQRCode
import csv
import datetime

# ═══════════════════════════════════════════════════════════════
# 1. ADMIN INDEX METRICS HOOK (PATCHING admin.site.index)
# ═══════════════════════════════════════════════════════════════
original_index = admin.site.index

def custom_admin_index(request, extra_context=None):
    extra_context = extra_context or {}
    today = timezone.now().date()
    today_start = timezone.make_aware(datetime.datetime.combine(today, datetime.time.min))
    
    # User counts
    users = User.objects.all()
    extra_context['total_users'] = users.count()
    extra_context['active_users'] = users.filter(is_active=True).count()
    extra_context['new_users_today'] = users.filter(date_joined__gte=today_start).count()
    
    # Subscriptions
    subs = Subscription.objects.all()
    extra_context['total_subs'] = subs.count()
    extra_context['active_subs'] = subs.filter(status='Active').count()
    extra_context['expired_subs'] = subs.filter(status='Expired').count()
    extra_context['pending_subs'] = subs.filter(status='Pending').count()
    extra_context['cancelled_subs'] = subs.filter(status='Cancelled').count()
    
    # Payments
    payments = Payment.objects.all()
    extra_context['total_payments'] = payments.count()
    extra_context['success_payments'] = payments.filter(payment_status='Paid').count()
    extra_context['failed_payments'] = payments.exclude(payment_status='Paid').count()
    
    # Revenue calculations
    thirty_days_ago = timezone.now() - datetime.timedelta(days=30)
    three_sixty_five_days_ago = timezone.now() - datetime.timedelta(days=365)
    extra_context['monthly_revenue'] = payments.filter(payment_status='Paid', created_at__gte=thirty_days_ago).aggregate(Sum('amount'))['amount__sum'] or 0
    extra_context['yearly_revenue'] = payments.filter(payment_status='Paid', created_at__gte=three_sixty_five_days_ago).aggregate(Sum('amount'))['amount__sum'] or 0
    
    # Assets
    extra_context['total_qrs'] = DynamicQRCode.objects.exclude(qr_type='custom-url').count()
    extra_context['total_shorts'] = DynamicQRCode.objects.filter(qr_type='custom-url').count()
    
    # Expiry list
    extra_context['expiring_today'] = subs.filter(status='Active', end_date__date=today)
    extra_context['expiring_3_days'] = subs.filter(status='Active', end_date__date__gt=today, end_date__date__lte=today + datetime.timedelta(days=3))
    extra_context['expiring_7_days'] = subs.filter(status='Active', end_date__date__gt=today + datetime.timedelta(days=3), end_date__date__lte=today + datetime.timedelta(days=7))
    extra_context['already_expired'] = subs.filter(status='Expired')
    
    return original_index(request, extra_context)

admin.site.index = custom_admin_index

# ═══════════════════════════════════════════════════════════════
# 2. GENERIC CSV EXPORT ACTIONS
# ═══════════════════════════════════════════════════════════════
def export_as_csv(modeladmin, request, queryset):
    """Universal model CSV exporting handler."""
    meta = modeladmin.model._meta
    field_names = [field.name for field in meta.fields]
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = f'attachment; filename={meta.verbose_name_plural}.csv'
    writer = csv.writer(response)
    writer.writerow(field_names)
    for obj in queryset:
        writer.writerow([getattr(obj, field) for field in field_names])
    return response

export_as_csv.short_description = "Export Selected Records to CSV"

# ═══════════════════════════════════════════════════════════════
# 3. PLAN ADMINISTRATION
# ═══════════════════════════════════════════════════════════════
@admin.register(Plan)
class PlanAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'monthly_price', 'yearly_price', 'max_dynamic_qrs', 'max_short_urls', 'created_at')
    search_fields = ('name', 'code')
    actions = [export_as_csv]

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

# ═══════════════════════════════════════════════════════════════
# 4. SUBSCRIPTION ADMINISTRATION (WITH DURATION MODIFIERS)
# ═══════════════════════════════════════════════════════════════
@admin.register(Subscription)
class SubscriptionAdmin(admin.ModelAdmin):
    list_display = ('user', 'plan', 'status', 'billing_cycle', 'payment_status', 'start_date', 'end_date')
    list_filter = ('plan', 'status', 'billing_cycle', 'payment_status')
    search_fields = ('user__username', 'user__email')
    actions = [
        'activate_subscription', 
        'deactivate_subscription', 
        'cancel_subscription', 
        'extend_one_month', 
        'extend_one_year', 
        export_as_csv
    ]

    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser

    # Actions
    def activate_subscription(self, request, queryset):
        queryset.update(status='Active', payment_status='Paid')
    activate_subscription.short_description = "Activate Subscription"

    def deactivate_subscription(self, request, queryset):
        queryset.update(status='Expired')
    deactivate_subscription.short_description = "Deactivate/Expire Subscription"

    def cancel_subscription(self, request, queryset):
        queryset.update(status='Cancelled')
    cancel_subscription.short_description = "Cancel Subscription"

    def extend_one_month(self, request, queryset):
        for sub in queryset:
            if sub.end_date:
                sub.end_date += datetime.timedelta(days=30)
                sub.save()
    extend_one_month.short_description = "Extend Expiry by 30 Days"

    def extend_one_year(self, request, queryset):
        for sub in queryset:
            if sub.end_date:
                sub.end_date += datetime.timedelta(days=365)
                sub.save()
    extend_one_year.short_description = "Extend Expiry by 365 Days"

# ═══════════════════════════════════════════════════════════════
# 5. PAYMENT INVOICE LOGS
# ═══════════════════════════════════════════════════════════════
@admin.register(Payment)
class PaymentAdmin(admin.ModelAdmin):
    list_display = ('transaction_id', 'user', 'subscription_plan', 'amount', 'payment_status', 'gateway', 'created_at')
    list_filter = ('payment_status', 'gateway', 'payment_mode')
    search_fields = ('transaction_id', 'user__username', 'user__email')
    readonly_fields = ('transaction_id',)
    actions = [export_as_csv]

    def subscription_plan(self, obj):
        return obj.subscription.plan.name if obj.subscription else '-'
    subscription_plan.short_description = "Plan"

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

# ═══════════════════════════════════════════════════════════════
# 6. SYSTEM EVENT LOGGER
# ═══════════════════════════════════════════════════════════════
@admin.register(ActivityLog)
class ActivityLogAdmin(admin.ModelAdmin):
    list_display = ('user', 'action', 'created_at', 'ip_address', 'status')
    list_filter = ('status', 'created_at')
    search_fields = ('user__username', 'action', 'ip_address')
    readonly_fields = ('user', 'action', 'created_at', 'ip_address', 'status')
    actions = [export_as_csv]

    def has_add_permission(self, request):
        return False

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

# ═══════════════════════════════════════════════════════════════
# 7. CUSTOM DETAILED USER PROFILE MONITORING & ANALYTICS
# ═══════════════════════════════════════════════════════════════
admin.site.unregister(User)

class RegistrationDateListFilter(admin.SimpleListFilter):
    title = 'Registration Date'
    parameter_name = 'reg_date'

    def lookups(self, request, modeladmin):
        return (
            ('today', "Today's Registrations"),
            ('month', "This Month's Registrations"),
        )

    def queryset(self, request, queryset):
        today = timezone.now().date()
        if self.value() == 'today':
            return queryset.filter(date_joined__date=today)
        if self.value() == 'month':
            return queryset.filter(date_joined__month=today.month, date_joined__year=today.year)

class UserSubscriptionListFilter(admin.SimpleListFilter):
    title = 'Subscription Status'
    parameter_name = 'sub_status'

    def lookups(self, request, modeladmin):
        return (
            ('active', 'Active Subscriptions'),
            ('expired', 'Expired Subscriptions'),
        )

    def queryset(self, request, queryset):
        if self.value() == 'active':
            return queryset.filter(subscriptions__status='Active').distinct()
        if self.value() == 'expired':
            return queryset.filter(subscriptions__status='Expired').distinct()

@admin.register(User)
class CustomUserAdmin(UserAdmin):
    list_display = (
        'username', 'email', 'get_plan_name', 'get_sub_status', 
        'get_remaining_days', 'get_total_paid', 'get_qr_count', 
        'get_short_count', 'is_active', 'date_joined'
    )
    list_filter = ('is_active', 'is_staff', RegistrationDateListFilter, UserSubscriptionListFilter)
    actions = [export_as_csv]
    
    # Append profile panels to the edit page fieldsets
    def get_fieldsets(self, request, obj=None):
        fieldsets = super().get_fieldsets(request, obj)
        if obj:
            fieldsets = list(fieldsets) + [
                ('SaaS Profiler & Analytics', {'fields': ('profile_summary_panel',)})
            ]
        return fieldsets

    def get_readonly_fields(self, request, obj=None):
        fields = super().get_readonly_fields(request, obj)
        if obj:
            fields = list(fields) + ['profile_summary_panel']
        return fields

    # Context counters
    def get_plan_name(self, obj):
        sub = Subscription.objects.filter(user=obj, status='Active').first()
        return sub.plan.name if sub else 'FREE'
    get_plan_name.short_description = "Plan"

    def get_sub_status(self, obj):
        sub = Subscription.objects.filter(user=obj, status='Active').first()
        return sub.status if sub else '-'
    get_sub_status.short_description = "Sub Status"

    def get_remaining_days(self, obj):
        sub = Subscription.objects.filter(user=obj, status='Active').first()
        if sub:
            if sub.plan.code == 'free':
                return '∞'
            if sub.end_date:
                delta = sub.end_date - timezone.now()
                return max(0, delta.days)
        return '-'
    get_remaining_days.short_description = "Days Left"

    def get_total_paid(self, obj):
        val = Payment.objects.filter(user=obj, payment_status='Paid').aggregate(Sum('amount'))['amount__sum'] or 0
        return f"₹{val}"
    get_total_paid.short_description = "Total Paid"

    def get_qr_count(self, obj):
        return DynamicQRCode.objects.filter(user=obj).exclude(qr_type='custom-url').count()
    get_qr_count.short_description = "QRs"

    def get_short_count(self, obj):
        return DynamicQRCode.objects.filter(user=obj, qr_type='custom-url').count()
    get_short_count.short_description = "Shorts"

    # HTML detailed user summary dashboard panel
    def profile_summary_panel(self, obj):
        sub = Subscription.objects.filter(user=obj, status='Active').first()
        payments = Payment.objects.filter(user=obj).order_by('-created_at')[:5]
        qrs = DynamicQRCode.objects.filter(user=obj).exclude(qr_type='custom-url')[:5]
        shorts = DynamicQRCode.objects.filter(user=obj, qr_type='custom-url')[:5]
        logs = ActivityLog.objects.filter(user=obj).order_by('-created_at')[:8]
        
        # Stats calculations
        qr_count = DynamicQRCode.objects.filter(user=obj).exclude(qr_type='custom-url').count()
        short_count = DynamicQRCode.objects.filter(user=obj, qr_type='custom-url').count()
        total_spent = Payment.objects.filter(user=obj, payment_status='Paid').aggregate(Sum('amount'))['amount__sum'] or 0
        
        plan_name = sub.plan.name if sub else 'FREE'
        cycle = sub.billing_cycle if sub else '-'
        sub_status = sub.status if sub else '-'
        sub_start = sub.start_date.strftime('%B %d, %Y') if sub else '-'
        sub_end = sub.end_date.strftime('%B %d, %Y') if sub and sub.end_date else ('Lifetime' if plan_name == 'FREE' else '-')
        
        # Max Limits counters
        max_qrs = sub.plan.max_dynamic_qrs if sub else 5
        max_shorts = sub.plan.max_short_urls if sub else 10
        
        # Build Profile summary
        html = f"""
        <div style="font-family: Arial, sans-serif; color: #333; margin-top: 10px;">
            <!-- Row 1: Profile & Limits -->
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px;">
                <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 15px;">
                    <h4 style="margin: 0 0 10px 0; color: #475569; font-weight: bold; border-bottom: 1px solid #e2e8f0; padding-bottom: 5px;">Active Subscription Detail</h4>
                    <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                        <tr><td style="padding: 4px 0; font-weight: bold; color: #64748b;">Plan Tier:</td><td style="text-align: right; font-weight: bold; color: #7c3aed;">{plan_name} ({cycle})</td></tr>
                        <tr><td style="padding: 4px 0; font-weight: bold; color: #64748b;">Status:</td><td style="text-align: right; font-weight: bold; color: #16a34a;">{sub_status}</td></tr>
                        <tr><td style="padding: 4px 0; font-weight: bold; color: #64748b;">Start Date:</td><td style="text-align: right;">{sub_start}</td></tr>
                        <tr><td style="padding: 4px 0; font-weight: bold; color: #64748b;">Expiry Date:</td><td style="text-align: right;">{sub_end}</td></tr>
                        <tr><td style="padding: 4px 0; font-weight: bold; color: #64748b;">Total Paid:</td><td style="text-align: right; font-weight: bold;">₹{total_spent}</td></tr>
                    </table>
                </div>
                
                <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; padding: 15px;">
                    <h4 style="margin: 0 0 10px 0; color: #475569; font-weight: bold; border-bottom: 1px solid #e2e8f0; padding-bottom: 5px;">Current Usage Stats & Limits</h4>
                    <div style="font-size: 12px; margin-bottom: 10px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 3px;">
                            <span>Dynamic QR Codes</span>
                            <span style="font-weight: bold;">{qr_count} / {max_qrs}</span>
                        </div>
                        <div style="height: 6px; background: #e2e8f0; border-radius: 3px; overflow: hidden;">
                            <div style="width: {min(100, (qr_count/max_qrs)*100)}%; height: 100%; background: #7c3aed;"></div>
                        </div>
                    </div>
                    <div style="font-size: 12px; margin-bottom: 10px;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 3px;">
                            <span>Short URL Redirects</span>
                            <span style="font-weight: bold;">{short_count} / {max_shorts}</span>
                        </div>
                        <div style="height: 6px; background: #e2e8f0; border-radius: 3px; overflow: hidden;">
                            <div style="width: {min(100, (short_count/max_shorts)*100)}%; height: 100%; background: #7c3aed;"></div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Row 2: Tables for history -->
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                <!-- Payments History -->
                <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 15px;">
                    <h4 style="margin: 0 0 10px 0; color: #475569; font-weight: bold; border-bottom: 1px solid #e2e8f0; padding-bottom: 5px;">Recent Invoices</h4>
                    <table style="width: 100%; font-size: 11px; text-align: left; border-collapse: collapse;">
                        <thead>
                            <tr style="color: #64748b; font-weight: bold; border-bottom: 1px solid #eee;">
                                <th style="padding: 4px 0;">TXN ID</th>
                                <th style="padding: 4px 0; text-align: right;">Amount</th>
                                <th style="padding: 4px 0; text-align: right;">Date</th>
                            </tr>
                        </thead>
                        <tbody>
        """
        
        for pay in payments:
            html += f"""
                            <tr style="border-bottom: 1px solid #f1f5f9;">
                                <td style="padding: 6px 0; font-family: monospace;">{pay.transaction_id}</td>
                                <td style="padding: 6px 0; text-align: right; font-weight: bold;">₹{pay.amount}</td>
                                <td style="padding: 6px 0; text-align: right; color: #64748b;">{pay.created_at.strftime('%Y-%m-%d')}</td>
                            </tr>
            """
        if not payments:
            html += """<tr><td colspan="3" style="text-align: center; color: #94a3b8; padding: 10px;">No transaction records</td></tr>"""
            
        html += """
                        </tbody>
                    </table>
                </div>

                <!-- Activity Log Timeline -->
                <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 15px;">
                    <h4 style="margin: 0 0 10px 0; color: #475569; font-weight: bold; border-bottom: 1px solid #e2e8f0; padding-bottom: 5px;">Recent Event Activity Timeline</h4>
                    <div style="max-height: 150px; overflow-y: auto; font-size: 11px; line-height: 1.5;">
        """
        
        for log in logs:
            html += f"""
                        <div style="padding: 4px 0; border-bottom: 1px solid #f8fafc; display: flex; justify-content: space-between;">
                            <span style="color: #334155; font-weight: 500;">{log.action}</span>
                            <span style="color: #94a3b8;">{log.created_at.strftime('%Y-%m-%d %H:%M')}</span>
                        </div>
            """
        if not logs:
            html += """<div style="text-align: center; color: #94a3b8; padding: 15px;">No logged events</div>"""
            
        html += """
                    </div>
                </div>
            </div>
        </div>
        """
        return mark_safe(html)
    
    profile_summary_panel.short_description = "User SaaS Profile Analytics Overview"
